from __future__ import annotations

import argparse
import csv
from datetime import datetime
import gc
import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import sys
import tempfile
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
OUTPUT_DIR = Path(__file__).resolve().parent
APP_PATH = ROOT / "app.py"
QUERY_MODULE_PATH = ROOT / "query_understanding.py"
ANSWER_CONTRACT_PATH = ROOT / "answer_contract.py"
DB_PATH = ROOT / "10_vector_db_with_major_accident_docs"
BASELINE_DIR = ROOT / "26_rag_retrieval_evaluation"
POSTFIX_DIR = ROOT / "27_rag_retrieval_evaluation_postfix"
BASELINE_SCRIPT = BASELINE_DIR / "run_rag_retrieval_quality_evaluation.py"
POSTFIX_SCRIPT = POSTFIX_DIR / "run_postfix_rag_retrieval_evaluation.py"
GOLD_PATH = BASELINE_DIR / "rag_retrieval_eval_questions_30.tsv"
DEVELOPMENT_PATH = OUTPUT_DIR / "development_question_matrix.jsonl"
HOLDOUT_PATH = OUTPUT_DIR / "holdout_question_matrix.jsonl"
DEVELOPMENT_RESULT_PATH = OUTPUT_DIR / "development_question_results.tsv"
HOLDOUT_RESULT_PATH = OUTPUT_DIR / "holdout_question_results.tsv"
SUMMARY_PATH = OUTPUT_DIR / "general_query_understanding_summary.json"
RETRIEVAL_RESULT_PATH = OUTPUT_DIR / "rag_retrieval_regression_results_30.tsv"
QUALITY_REPORT_PATH = OUTPUT_DIR / "general_query_understanding_quality_report.txt"
MANIFEST_PATH = OUTPUT_DIR / "evaluation_manifest.json"

PREVIOUS_APP_SHA256 = "3D798DA421F8BBE08B16135C85E084ACB85D0ACD16AC8939A647A184A27E4680"
GIT_HEAD_BEFORE_WORK = "8bb9828eb9fc65230c531925faf2499d34ce325b"
DEVELOPMENT_SHA256 = "C7EBF7D857E128787D0034D1EE171EE5C8FA1734AA7868479E0E102CCDBB165F"
HOLDOUT_SHA256 = "52F1C4D2A4D662EC4CFF28976220F9580752B113C467E9D2EEC968EB4B85B9F2"
GOLD_SHA256 = "F950A1505743A4B3607F0EC3BC60D8ED477FB6C3627F14FCD156BF7C81B7D076"
COLLECTION_NAME = "mine_safety_docs"
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
EMBEDDING_DIMENSION = 384
TOP_K = 5
MINIMUM_RETRIEVAL_METRICS = {
    "document_hit_at_3": 0.5333,
    "document_hit_at_5": 0.5667,
    "mrr": 0.4639,
    "metadata_completeness": 0.7500,
}
NOT_FINAL_NOTICE = (
    "개발·보류 확인 및 기존 30문항 재사용 회귀평가이며 교수님 최종 5문항을 사용한 "
    "최종 잠금 평가가 아니다. 이 결과만으로 일반화 성능·답변 정확성·법적 정확도를 확정하지 않는다."
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest().upper()


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    return [
        json.loads(line)
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]


def read_tsv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle, delimiter="\t"))


def write_tsv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, delimiter="\t", lineterminator="\n", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"모듈을 읽을 수 없습니다: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def physical_db_snapshot(path: Path) -> dict[str, Any]:
    records: list[dict[str, Any]] = []
    protected_names = {"data_level0.bin", "length.bin", "chroma.sqlite3"}
    for file_path in sorted(candidate for candidate in path.rglob("*") if candidate.is_file()):
        stat = file_path.stat()
        relative = file_path.relative_to(path).as_posix()
        file_hash = ""
        hash_error = ""
        if file_path.name in protected_names:
            try:
                file_hash = sha256_file(file_path)
            except OSError as error:
                hash_error = f"{type(error).__name__}: {error}"
        records.append(
            {
                "path": relative,
                "size": stat.st_size,
                "mtime_ns": stat.st_mtime_ns,
                "sha256": file_hash,
                "sha256_error": hash_error,
            }
        )
    metadata = "\n".join(
        f"{row['path']}\t{row['size']}\t{row['mtime_ns']}\t{row['sha256']}"
        for row in records
    )
    return {
        "file_count": len(records),
        "total_bytes": sum(row["size"] for row in records),
        "metadata_sha256": sha256_text(metadata),
        "protected_files": [row for row in records if Path(row["path"]).name in protected_names],
    }


def copy_db_to_temporary() -> tuple[Path, Path]:
    temporary_root = Path(tempfile.mkdtemp(prefix="MineSafe_RAG_General_Query_Eval_"))
    temp_base = Path(tempfile.gettempdir()).resolve()
    if temp_base not in temporary_root.resolve().parents:
        raise RuntimeError("임시 DB 경로가 Windows 임시 폴더 밖에 생성됐습니다.")
    copied_db = temporary_root / DB_PATH.name
    shutil.copytree(DB_PATH, copied_db, copy_function=shutil.copy2)
    return temporary_root, copied_db


def cleanup_temporary(path: Path) -> None:
    temp_base = Path(tempfile.gettempdir()).resolve()
    resolved = path.resolve()
    if temp_base not in resolved.parents or not path.name.startswith("MineSafe_RAG_General_Query_Eval_"):
        raise RuntimeError(f"삭제를 거부한 임시 경로: {resolved}")
    shutil.rmtree(resolved)


def matrix_metrics(rows: list[dict[str, Any]], analyze_query) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    results: list[dict[str, Any]] = []
    for row in rows:
        analysis = analyze_query(row["question"])
        expected_secondary = set(row["expected_secondary_domains"])
        predicted_secondary = set(analysis.secondary_domains)
        required_signals = set(row["required_signals"])
        predicted_signals = set(analysis.hazard_signals)
        required_outputs = set(row["required_requested_outputs"])
        predicted_outputs = set(analysis.requested_outputs)
        results.append(
            {
                "case_id": row["case_id"],
                "split": row["split"],
                "question": row["question"],
                "expected_primary_domain": row["expected_primary_domain"],
                "predicted_primary_domain": analysis.primary_domain,
                "primary_correct": int(analysis.primary_domain == row["expected_primary_domain"]),
                "expected_secondary_domains": json.dumps(row["expected_secondary_domains"], ensure_ascii=False),
                "predicted_secondary_domains": json.dumps(analysis.secondary_domains, ensure_ascii=False),
                "secondary_hit_count": len(expected_secondary & predicted_secondary),
                "secondary_expected_count": len(expected_secondary),
                "required_signals": json.dumps(row["required_signals"], ensure_ascii=False),
                "predicted_signals": json.dumps(analysis.hazard_signals, ensure_ascii=False),
                "signal_hit_count": len(required_signals & predicted_signals),
                "signal_expected_count": len(required_signals),
                "expected_stage": row["expected_stage"],
                "predicted_stage": analysis.work_stage,
                "stage_correct": int(analysis.work_stage == row["expected_stage"]),
                "required_requested_outputs": json.dumps(row["required_requested_outputs"], ensure_ascii=False),
                "predicted_requested_outputs": json.dumps(analysis.requested_outputs, ensure_ascii=False),
                "requested_output_hit_count": len(required_outputs & predicted_outputs),
                "requested_output_expected_count": len(required_outputs),
                "expected_should_clarify": int(bool(row["should_clarify"])),
                "predicted_should_clarify": int(bool(analysis.clarification_question)),
                "ambiguity_correct": int(bool(analysis.clarification_question) == bool(row["should_clarify"])),
                "clarification_question": analysis.clarification_question,
                "search_query_count": len(analysis.search_queries),
                "search_queries": json.dumps(analysis.search_queries, ensure_ascii=False),
                "forbidden_dominant_topics": json.dumps(row["forbidden_dominant_topics"], ensure_ascii=False),
                "note": row["note"],
            }
        )

    total = len(results)
    secondary_expected = sum(row["secondary_expected_count"] for row in results)
    signal_expected = sum(row["signal_expected_count"] for row in results)
    output_expected = sum(row["requested_output_expected_count"] for row in results)
    metrics = {
        "question_count": total,
        "primary_domain_accuracy": sum(row["primary_correct"] for row in results) / total,
        "secondary_domain_recall": (
            sum(row["secondary_hit_count"] for row in results) / secondary_expected
            if secondary_expected else 1.0
        ),
        "hazard_signal_coverage": sum(row["signal_hit_count"] for row in results) / signal_expected,
        "work_stage_accuracy": sum(row["stage_correct"] for row in results) / total,
        "requested_output_detection_recall": sum(row["requested_output_hit_count"] for row in results) / output_expected,
        "ambiguity_handling_accuracy": sum(row["ambiguity_correct"] for row in results) / total,
        "major_safety_misclassification_count": sum(not row["primary_correct"] for row in results),
        "major_safety_misclassification_ids": [row["case_id"] for row in results if not row["primary_correct"]],
        "average_search_query_count": sum(row["search_query_count"] for row in results) / total,
        "maximum_search_query_count": max(row["search_query_count"] for row in results),
    }
    return metrics, results


def retrieve_current_top5(
    question: str,
    collection: Any,
    model: Any,
    namespace: dict[str, Any],
    top_k: int = TOP_K,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    from query_understanding import analyze_query, rerank_results_by_understanding

    understanding = analyze_query(question)
    context = namespace["extract_query_context_signals"](question)
    intent = namespace["detect_question_intent"](context["normalized_question"])
    expanded = namespace["expand_search_query"](question, intent)
    search_queries = list(dict.fromkeys(
        query.strip()
        for query in [expanded, *understanding.search_queries]
        if query and query.strip()
    ))[:4]
    embeddings = model.encode(search_queries, normalize_embeddings=True, show_progress_bar=False).tolist()
    if not embeddings or any(len(row) != EMBEDDING_DIMENSION for row in embeddings):
        raise RuntimeError("질의 embedding 차원이 384가 아닙니다.")

    question_type = namespace["classify_question_type"](expanded)
    rerank_config = namespace["RERANK_TYPE_CONFIGS"].get(question_type)
    internal_count = top_k
    if rerank_config:
        internal_count = max(top_k, int(rerank_config["internal_count"]))
    elif context.get("detail_signal") in {"rainy_pre_blasting", "roof_water_increase"}:
        internal_count = max(top_k, 20)
    internal_count = min(internal_count, int(collection.count()))
    raw = collection.query(
        query_embeddings=embeddings,
        n_results=internal_count,
        include=["documents", "metadatas", "distances"],
    )
    candidates: list[dict[str, Any]] = []
    seen: set[str] = set()
    for query_index, search_query in enumerate(search_queries):
        documents = raw.get("documents", [[]])[query_index]
        metadatas = raw.get("metadatas", [[]])[query_index]
        distances = raw.get("distances", [[]])[query_index]
        for index, document in enumerate(documents):
            metadata = metadatas[index] if index < len(metadatas) and isinstance(metadatas[index], dict) else {}
            distance = distances[index] if index < len(distances) else None
            candidate = namespace["make_search_result"](
                document or "",
                metadata,
                distance,
                vector_rank=(query_index * internal_count) + index + 1,
            )
            key = str(candidate.get("chunk_id", "")).strip() or f"{candidate.get('source', '')}|{candidate.get('text', '')[:160]}"
            if key in seen:
                continue
            seen.add(key)
            candidate["matched_search_queries"] = [search_query]
            candidates.append(candidate)
    if rerank_config:
        candidates = namespace["add_type_source_candidates"](
            collection,
            embeddings[0],
            candidates,
            rerank_config["preferred_files"],
        )
    candidate_pool = namespace["rerank_search_results"](
        expanded,
        candidates,
        max(top_k, min(20, len(candidates))),
    )
    selected = rerank_results_by_understanding(understanding, candidate_pool, top_k)
    return selected, {
        "question_intent": intent,
        "question_type": question_type,
        "expanded_search_query": expanded,
        "search_queries": search_queries,
        "search_query_count": len(search_queries),
        "internal_count": internal_count,
        "reranking_applied": bool(selected and selected[0].get("reranking_applied", False)),
        "detail_signal": context.get("detail_signal", ""),
        "question_context_label": context.get("display_label", ""),
    }


def evaluate_retrieval(copied_db: Path) -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    from answer_contract import build_complete_core_judgment
    from query_understanding import analyze_query

    base = load_module("minesafe_general_eval_base", BASELINE_SCRIPT)
    postfix = load_module("minesafe_general_eval_postfix", POSTFIX_SCRIPT)
    base.OPERATIONAL_FUNCTIONS = set(base.OPERATIONAL_FUNCTIONS) | {"extract_query_context_signals"}
    namespace = base.load_operational_search_namespace(APP_PATH)
    namespace["analyze_query"] = analyze_query
    namespace["build_complete_core_judgment"] = build_complete_core_judgment
    base.validate_operational_contract(namespace)
    postfix.retrieve_current_operational_top5 = retrieve_current_top5

    import chromadb

    client = chromadb.PersistentClient(path=str(copied_db))
    collection = client.get_collection(name=COLLECTION_NAME)
    logical_before = base.snapshot_collection_logical(collection)
    if logical_before["count"] != 820:
        raise RuntimeError(f"임시 DB collection count가 820이 아닙니다: {logical_before['count']}")
    model = base.load_local_embedding_model(EMBEDDING_MODEL_NAME)
    gold_rows = read_tsv(GOLD_PATH)
    result_rows, _detail_rows, _element_rows, failure_ids = postfix.evaluate_questions(
        base,
        gold_rows,
        collection,
        model,
        namespace,
    )
    metrics = postfix.metrics_from_rows(result_rows)
    logical_after = base.snapshot_collection_logical(collection)
    if logical_before != logical_after:
        raise RuntimeError("임시 DB의 논리 스냅샷이 평가 전후 달라졌습니다.")
    client.close()
    del collection, client, model
    gc.collect()
    return metrics, result_rows, {
        "before": logical_before,
        "after": logical_after,
        "search_failure_ids": failure_ids,
    }


def validate_targets(development: dict[str, Any], holdout: dict[str, Any], retrieval: dict[str, Any]) -> None:
    for split, score in (("development", development), ("holdout", holdout)):
        if score["primary_domain_accuracy"] < 0.93:
            raise RuntimeError(f"{split} primary domain 정확도 목표 미달")
        if score["work_stage_accuracy"] < 0.90:
            raise RuntimeError(f"{split} 작업 단계 정확도 목표 미달")
        if score["requested_output_detection_recall"] < 0.95:
            raise RuntimeError(f"{split} 요청항목 recall 목표 미달")
        if score["major_safety_misclassification_count"]:
            raise RuntimeError(f"{split} 중대한 안전 오분류 발생")
        if score["maximum_search_query_count"] > 4:
            raise RuntimeError(f"{split} 검색 질의 수가 4개를 초과")
    for field, minimum in MINIMUM_RETRIEVAL_METRICS.items():
        if round(float(retrieval[field]), 4) < minimum:
            raise RuntimeError(f"30문항 검색 회귀 지표 하락: {field}={retrieval[field]:.4f} < {minimum:.4f}")
    if int(retrieval["search_failure_count"]):
        raise RuntimeError("30문항 검색 실행 실패가 발생했습니다.")


def result_columns(rows: list[dict[str, Any]]) -> list[str]:
    return list(rows[0]) if rows else []


def run() -> dict[str, Any]:
    if sha256_file(DEVELOPMENT_PATH) != DEVELOPMENT_SHA256:
        raise RuntimeError("고정 개발 질문 세트 SHA-256이 달라졌습니다.")
    if sha256_file(HOLDOUT_PATH) != HOLDOUT_SHA256:
        raise RuntimeError("고정 보류 질문 세트 SHA-256이 달라졌습니다.")
    if sha256_file(GOLD_PATH) != GOLD_SHA256:
        raise RuntimeError("기존 30문항 Gold SHA-256이 달라졌습니다.")

    from query_understanding import analyze_query

    development_rows = read_jsonl(DEVELOPMENT_PATH)
    holdout_rows = read_jsonl(HOLDOUT_PATH)
    development_metrics, development_results = matrix_metrics(development_rows, analyze_query)
    holdout_metrics, holdout_results = matrix_metrics(holdout_rows, analyze_query)
    write_tsv(DEVELOPMENT_RESULT_PATH, development_results, result_columns(development_results))
    write_tsv(HOLDOUT_RESULT_PATH, holdout_results, result_columns(holdout_results))

    original_db_before = physical_db_snapshot(DB_PATH)
    temporary_root: Path | None = None
    try:
        temporary_root, copied_db = copy_db_to_temporary()
        retrieval_metrics, retrieval_results, logical_snapshot = evaluate_retrieval(copied_db)
    finally:
        if temporary_root is not None and temporary_root.exists():
            cleanup_temporary(temporary_root)
    original_db_after = physical_db_snapshot(DB_PATH)
    if original_db_before != original_db_after:
        raise RuntimeError("원본 Vector DB의 물리 상태가 평가 전후 달라졌습니다.")

    write_tsv(RETRIEVAL_RESULT_PATH, retrieval_results, result_columns(retrieval_results))
    validate_targets(development_metrics, holdout_metrics, retrieval_metrics)

    outside_questions = ["오늘 저녁 메뉴 추천", "주식 종목 추천", "주말 여행지", "야구 경기 결과"]
    outside_failures = [
        question for question in outside_questions
        if analyze_query(question).in_scope
    ]
    if outside_failures:
        raise RuntimeError(f"범위 밖 질문 차단 실패: {outside_failures}")

    source = QUERY_MODULE_PATH.read_text(encoding="utf-8")
    hardcoded_questions = [
        row["question"] for row in [*development_rows, *holdout_rows]
        if row["question"] in source
    ]
    if hardcoded_questions:
        raise RuntimeError("개발·보류 질문 전체 문장이 구현 코드에 포함됐습니다.")

    summary = {
        "evaluation_name": "범용 질문 이해 및 답변 계약 개발·보류 확인",
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "development_metrics": development_metrics,
        "holdout_metrics": holdout_metrics,
        "scope_block_accuracy": (len(outside_questions) - len(outside_failures)) / len(outside_questions),
        "scope_block_failures": outside_failures,
        "exact_full_sentence_hardcoding_count": len(hardcoded_questions),
        "retrieval_regression_metrics": retrieval_metrics,
        "retrieval_minimums": MINIMUM_RETRIEVAL_METRICS,
        "temporary_vector_db_used": True,
        "original_vector_db_unchanged": original_db_before == original_db_after,
        "original_vector_db_before": original_db_before,
        "original_vector_db_after": original_db_after,
        "temporary_vector_db_logical_snapshot": logical_snapshot,
        "not_final_lock_notice": NOT_FINAL_NOTICE,
    }
    SUMMARY_PATH.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    report_lines = [
        "MineSafe AI 범용 질문 이해 및 답변 계약 품질 보고",
        "",
        f"생성시각: {summary['created_at']}",
        NOT_FINAL_NOTICE,
        "",
        "[개발 질문]",
        json.dumps(development_metrics, ensure_ascii=False, indent=2),
        "",
        "[보류 확인 질문]",
        json.dumps(holdout_metrics, ensure_ascii=False, indent=2),
        "",
        "[범위 밖 질문 차단]",
        f"정확도: {summary['scope_block_accuracy']:.4f}",
        f"실패: {outside_failures}",
        "",
        "[전체 문장 하드코딩 검사]",
        f"발견 수: {len(hardcoded_questions)}",
        "",
        "[기존 30문항 검색 회귀]",
        json.dumps(retrieval_metrics, ensure_ascii=False, indent=2),
        "",
        "[Vector DB 보호]",
        "원본 DB는 ChromaDB client로 열지 않았고 Windows 임시 폴더의 전체 복사본만 평가에 사용했다.",
        f"원본 물리 상태 전후 동일: {original_db_before == original_db_after}",
        f"임시 DB 논리 스냅샷 전후 동일: {logical_snapshot['before'] == logical_snapshot['after']}",
        "",
        "외부 API·인터넷 호출: 0",
        "교수님 최종 5문항 사용: 0",
    ]
    QUALITY_REPORT_PATH.write_text("\n".join(report_lines) + "\n", encoding="utf-8")

    gold_rows = read_tsv(GOLD_PATH)
    gold_question_sha = sha256_text("\n".join(f"{row['eval_id']}\t{row['question']}" for row in gold_rows))
    result_hashes = {
        path.name: sha256_file(path)
        for path in (
            DEVELOPMENT_RESULT_PATH,
            HOLDOUT_RESULT_PATH,
            SUMMARY_PATH,
            RETRIEVAL_RESULT_PATH,
            QUALITY_REPORT_PATH,
        )
    }
    manifest = {
        "evaluation_name": "범용 질문 이해 및 답변 계약 개발·보류·검색 회귀평가",
        "evaluation_type": "development and holdout query-understanding evaluation with retrieval regression",
        "created_at": summary["created_at"],
        "previous_app_sha256": PREVIOUS_APP_SHA256,
        "current_app_sha256": sha256_file(APP_PATH),
        "git_head_before_work": GIT_HEAD_BEFORE_WORK,
        "development_question_sha256": DEVELOPMENT_SHA256,
        "holdout_question_sha256": HOLDOUT_SHA256,
        "existing_30_gold_sha256": GOLD_SHA256,
        "existing_30_question_sha256": gold_question_sha,
        "development_question_count": len(development_rows),
        "holdout_question_count": len(holdout_rows),
        "development_metrics": development_metrics,
        "holdout_metrics": holdout_metrics,
        "retrieval_regression_metrics_30": retrieval_metrics,
        "evaluation_script_sha256": sha256_file(Path(__file__)),
        "query_understanding_module_sha256": sha256_file(QUERY_MODULE_PATH),
        "answer_contract_module_sha256": sha256_file(ANSWER_CONTRACT_PATH),
        "result_files_sha256": result_hashes,
        "python_executable": sys.executable,
        "embedding_model": EMBEDDING_MODEL_NAME,
        "embedding_dimension": EMBEDDING_DIMENSION,
        "vector_db_source": "Windows 임시 폴더에 만든 10_vector_db_with_major_accident_docs 전체 복사본",
        "collection_name": COLLECTION_NAME,
        "collection_count": logical_snapshot["before"]["count"],
        "external_api_or_internet_used": False,
        "original_vector_db_unchanged": original_db_before == original_db_after,
        "not_final_lock_notice": NOT_FINAL_NOTICE,
    }
    MANIFEST_PATH.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.parse_args()
    manifest = run()
    print(json.dumps({
        "development": manifest["development_metrics"],
        "holdout": manifest["holdout_metrics"],
        "retrieval": manifest["retrieval_regression_metrics_30"],
        "original_vector_db_unchanged": manifest["original_vector_db_unchanged"],
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
