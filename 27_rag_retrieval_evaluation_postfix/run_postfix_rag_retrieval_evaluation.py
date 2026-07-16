from __future__ import annotations

import argparse
import csv
from collections import Counter
from datetime import datetime
import gc
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Any


ROOT_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT_DIR / "27_rag_retrieval_evaluation_postfix"
BASELINE_DIR = ROOT_DIR / "26_rag_retrieval_evaluation"
BASELINE_SCRIPT_PATH = BASELINE_DIR / "run_rag_retrieval_quality_evaluation.py"
SOURCE_QUESTIONS_PATH = ROOT_DIR / "02_질문시나리오" / "question_scenarios_110.tsv"
APP_PATH = ROOT_DIR / "app.py"
DB_PATH = ROOT_DIR / "10_vector_db_with_major_accident_docs"
GOLD_PATH = BASELINE_DIR / "rag_retrieval_eval_questions_30.tsv"
BASELINE_RESULT_PATH = BASELINE_DIR / "rag_retrieval_eval_results_30.tsv"
BUILDER_PATH = OUTPUT_DIR / "build_postfix_workbooks.mjs"

RESULT_TSV_PATH = OUTPUT_DIR / "rag_retrieval_postfix_results_30.tsv"
RESULT_XLSX_PATH = OUTPUT_DIR / "rag_retrieval_postfix_results_30.xlsx"
MANUAL_XLSX_PATH = OUTPUT_DIR / "rag_retrieval_postfix_manual_review_30.xlsx"
SUMMARY_XLSX_PATH = OUTPUT_DIR / "rag_retrieval_postfix_summary.xlsx"
COMPARISON_XLSX_PATH = OUTPUT_DIR / "rag_retrieval_baseline_vs_postfix.xlsx"
COMPARISON_PNG_PATH = OUTPUT_DIR / "rag_retrieval_baseline_vs_postfix.png"
QUALITY_REPORT_PATH = OUTPUT_DIR / "rag_retrieval_postfix_quality_report.txt"
MANIFEST_PATH = OUTPUT_DIR / "evaluation_manifest.json"

COLLECTION_NAME = "mine_safety_docs"
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
EMBEDDING_DIMENSION = 384
NORMALIZE_EMBEDDINGS = True
TOP_K = 5
QUESTION_COUNT = 30

BASELINE_APP_SHA256 = "983B80A2C390829BA7750830696BC333E8AEABC1A7A4F7407237995AB4AC2FF2"
POSTFIX_APP_SHA256 = "3D798DA421F8BBE08B16135C85E084ACB85D0ACD16AC8939A647A184A27E4680"
ORIGINAL_QUESTION_SET_SHA256 = "544888A7717DD31CB0AF5099128D8493DE7AD50EB38F9ACE90FEFA2AFD72277A"
GOLD_SET_SHA256 = "F950A1505743A4B3607F0EC3BC60D8ED477FB6C3627F14FCD156BF7C81B7D076"

EVALUATION_NAME = "질문 context 및 미완성 답변 안전장치 수정 후 회귀평가"
EVALUATION_TYPE = "post-fix regression evaluation"
NOT_FINAL_LOCK_NOTICE = (
    "동일한 30문항을 재사용한 기존 기준선과의 변화 확인용 수정 후 회귀평가이며, "
    "최종 잠금 평가·독립 평가·완전한 블라인드 평가가 아니다."
)
LEGAL_ACCURACY_NOTICE = (
    "RAG 검색 지표는 법적 정확도, 최종 답변 정확도 또는 일반화 성능 확정을 직접 의미하지 않는다."
)

BASELINE_FILE_SHA256 = {
    "rag_retrieval_category_hit3.png": "7666472B8DCA36D21966736992AFF14FB3BEEDFAF86B01EA09C3C53F026D691B",
    "rag_retrieval_difficulty_hit3.png": "D9F03B943AC53832574FAC0B05AA29798862C8F248AD254AF93257E3CD3A6236",
    "rag_retrieval_eval_questions_30.tsv": "F950A1505743A4B3607F0EC3BC60D8ED477FB6C3627F14FCD156BF7C81B7D076",
    "rag_retrieval_eval_results_30.tsv": "68C0862FE0EA14BAD8985EB98151600059928DA3951A4C24B543682D1FB194CE",
    "rag_retrieval_eval_results_30.xlsx": "12B14788347755ADD84892ABA02A07C748A4AF4B1CC4BA4497CAB7D6A4FBB030",
    "rag_retrieval_hit_at_k.png": "6C04B014AB6C9695E5BA19D346B56C0C89E2607E14F524053A41D56A04277154",
    "rag_retrieval_manual_review_30.xlsx": "48DE0F14C00D66272847B12C377885AD9F0BA4035178F6324262E64A50485FA4",
    "rag_retrieval_quality_report.txt": "6533447015B081CDD78D3603716047E82B3240FD521B7D255E5C7E96EFB4DFD1",
    "rag_retrieval_summary.xlsx": "CCDCD08400CB91A17093A75BF9CF410E71CFF5442C09428D75C303E380EC9E2C",
    "run_rag_retrieval_quality_evaluation.py": "E1390BE2038384B24A53387AA1DAB648909958772266BD5AE50D20432CA92142",
}

EXPECTED_BASELINE_METRICS = {
    "document_hit_at_1": 0.4000,
    "document_hit_at_3": 0.5333,
    "document_hit_at_5": 0.5667,
    "mrr": 0.4639,
    "element_coverage_at_1": 0.1135,
    "element_coverage_at_3": 0.2294,
    "element_coverage_at_5": 0.2856,
    "metadata_completeness": 0.7500,
    "manual_review_count": 28,
    "search_failure_count": 0,
}

RATE_METRIC_LABELS = {
    "document_hit_at_1": "Document Hit@1",
    "document_hit_at_3": "Document Hit@3",
    "document_hit_at_5": "Document Hit@5",
    "mrr": "MRR",
    "element_coverage_at_1": "Element Coverage@1",
    "element_coverage_at_3": "Element Coverage@3",
    "element_coverage_at_5": "Element Coverage@5",
    "metadata_completeness": "Metadata 완전성",
}
COUNT_METRIC_LABELS = {
    "manual_review_count": "수동 검토 필요 문항 수",
    "search_failure_count": "검색 실행 실패 문항 수",
}


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def read_tsv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle, delimiter="\t"))


def load_baseline_module():
    spec = importlib.util.spec_from_file_location(
        "minesafe_baseline_rag_retrieval_eval",
        BASELINE_SCRIPT_PATH,
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("기존 평가 스크립트 모듈 사양을 만들 수 없습니다.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def assert_baseline_files_unchanged() -> dict[str, str]:
    current = {
        name: sha256_file(BASELINE_DIR / name)
        for name in BASELINE_FILE_SHA256
    }
    mismatches = {
        name: {"expected": BASELINE_FILE_SHA256[name], "actual": current[name]}
        for name in BASELINE_FILE_SHA256
        if current[name] != BASELINE_FILE_SHA256[name]
    }
    if mismatches:
        raise RuntimeError(
            "기존 26번 평가 기준선 파일 SHA-256 불일치: "
            + json.dumps(mismatches, ensure_ascii=False)
        )
    return current


def metrics_from_rows(rows: list[dict[str, Any]]) -> dict[str, float | int]:
    if not rows:
        raise RuntimeError("평가 결과 행이 비어 있습니다.")

    def mean(field: str) -> float:
        return sum(float(row[field]) for row in rows) / len(rows)

    return {
        "document_hit_at_1": mean("document_hit_at_1"),
        "document_hit_at_3": mean("document_hit_at_3"),
        "document_hit_at_5": mean("document_hit_at_5"),
        "mrr": mean("reciprocal_rank"),
        "element_coverage_at_1": mean("element_coverage_at_1"),
        "element_coverage_at_3": mean("element_coverage_at_3"),
        "element_coverage_at_5": mean("element_coverage_at_5"),
        "metadata_completeness": mean("metadata_completeness_top5"),
        "manual_review_count": sum(row["auto_review_status"] == "REVIEW" for row in rows),
        "search_failure_count": sum(bool(str(row.get("search_error", "")).strip()) for row in rows),
    }


def validate_baseline_metrics(metrics: dict[str, float | int]) -> None:
    for key, expected in EXPECTED_BASELINE_METRICS.items():
        actual = metrics[key]
        if isinstance(expected, int):
            if int(actual) != expected:
                raise RuntimeError(f"기존 기준선 지표 불일치: {key}={actual}, expected={expected}")
        elif round(float(actual), 4) != expected:
            raise RuntimeError(f"기존 기준선 지표 불일치: {key}={actual}, expected={expected}")


def retrieve_current_operational_top5(
    question: str,
    collection: Any,
    model: Any,
    namespace: dict[str, Any],
    top_k: int = TOP_K,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    if top_k != TOP_K:
        raise RuntimeError(f"평가 Top K는 {TOP_K}로 고정되어 있습니다.")
    question_context = namespace["extract_query_context_signals"](question)
    intent = namespace["detect_question_intent"](question_context["normalized_question"])
    expanded_question = namespace["expand_search_query"](question, intent)
    encoded = model.encode(
        [expanded_question],
        normalize_embeddings=NORMALIZE_EMBEDDINGS,
        show_progress_bar=False,
    )
    query_embedding = encoded.tolist()
    if not query_embedding or len(query_embedding[0]) != EMBEDDING_DIMENSION:
        raise RuntimeError("질의 embedding 차원이 384가 아닙니다.")

    question_type = namespace["classify_question_type"](expanded_question)
    rerank_config = namespace["RERANK_TYPE_CONFIGS"].get(question_type)
    internal_count = top_k
    if rerank_config:
        internal_count = max(top_k, int(rerank_config["internal_count"]))
    elif question_context.get("detail_signal") in {
        "rainy_pre_blasting",
        "roof_water_increase",
    }:
        internal_count = max(top_k, 20)
    internal_count = min(internal_count, int(collection.count()))

    raw = collection.query(
        query_embeddings=query_embedding,
        n_results=internal_count,
        include=["documents", "metadatas", "distances"],
    )
    documents = raw.get("documents", [[]])[0]
    metadatas = raw.get("metadatas", [[]])[0]
    distances = raw.get("distances", [[]])[0]
    candidates = []
    for index, document in enumerate(documents):
        metadata = (
            metadatas[index]
            if index < len(metadatas) and isinstance(metadatas[index], dict)
            else {}
        )
        distance = distances[index] if index < len(distances) else None
        candidates.append(
            namespace["make_search_result"](
                document or "",
                metadata,
                distance,
                vector_rank=index + 1,
            )
        )
    if rerank_config:
        candidates = namespace["add_type_source_candidates"](
            collection,
            query_embedding[0],
            candidates,
            rerank_config["preferred_files"],
        )
    selected = namespace["rerank_search_results"](
        expanded_question,
        candidates,
        top_k,
    )
    return selected, {
        "question_intent": intent,
        "question_type": question_type,
        "expanded_search_query": expanded_question,
        "internal_count": internal_count,
        "reranking_applied": bool(
            selected and selected[0].get("reranking_applied", False)
        ),
        "detail_signal": question_context.get("detail_signal", ""),
        "question_context_label": question_context.get("display_label", ""),
    }


def evaluate_questions(
    base: Any,
    gold_rows: list[dict[str, str]],
    collection: Any,
    model: Any,
    namespace: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    result_rows: list[dict[str, Any]] = []
    detail_rows: list[dict[str, Any]] = []
    element_rows: list[dict[str, Any]] = []
    failure_ids: list[str] = []

    for gold in gold_rows:
        patterns = json.loads(gold["expected_document_patterns"])
        core_terms = json.loads(gold["expected_core_terms"])
        search_error = ""
        diagnostics: dict[str, Any] = {}
        try:
            raw_hits, diagnostics = retrieve_current_operational_top5(
                gold["question"],
                collection,
                model,
                namespace,
                TOP_K,
            )
            hits = [
                base.enrich_retrieval_hit(hit, patterns, core_terms)
                for hit in raw_hits
            ]
        except Exception as error:
            search_error = f"{type(error).__name__}: {str(error)[:200]}"
            failure_ids.append(gold["eval_id"])
            hits = []

        document_metrics = base.compute_document_hit_metrics(hits)
        coverage_1, matched_1, reasons_1 = base.compute_element_coverage(hits, core_terms, 1)
        coverage_3, matched_3, reasons_3 = base.compute_element_coverage(hits, core_terms, 3)
        coverage_5, matched_5, reasons_5 = base.compute_element_coverage(hits, core_terms, 5)
        diversity = base.compute_retrieval_diversity(hits)
        metadata_completeness = base.compute_metadata_completeness(hits)
        metrics = {
            **document_metrics,
            "element_coverage_at_1": coverage_1,
            "element_coverage_at_3": coverage_3,
            "element_coverage_at_5": coverage_5,
            "metadata_completeness_top5": metadata_completeness,
        }
        auto_status, auto_reason = base.decide_auto_review(
            patterns,
            metrics,
            diversity,
            search_error,
        )
        result: dict[str, Any] = {
            "eval_id": gold["eval_id"],
            "source_question_id": gold["source_question_id"],
            "category": gold["normalized_category"],
            "difficulty": gold["difficulty"],
            "question": gold["question"],
            "expected_document_raw": gold["expected_document_raw"],
            "expected_document_patterns": gold["expected_document_patterns"],
            "expected_core_elements_raw": gold["expected_core_elements_raw"],
            "expected_core_terms": gold["expected_core_terms"],
            "gold_set_sha256": GOLD_SET_SHA256,
            **document_metrics,
            "element_coverage_at_1": coverage_1,
            "element_coverage_at_3": coverage_3,
            "element_coverage_at_5": coverage_5,
            **diversity,
            "metadata_completeness_top5": metadata_completeness,
            "auto_review_status": auto_status,
            "auto_review_reason": auto_reason,
            "question_intent": diagnostics.get("question_intent", ""),
            "question_type": diagnostics.get("question_type", ""),
            "expanded_search_query": diagnostics.get("expanded_search_query", ""),
            "reranking_applied": diagnostics.get("reranking_applied", False),
            "search_error": search_error,
        }
        for rank in range(1, TOP_K + 1):
            hit = hits[rank - 1] if rank <= len(hits) else {}
            result[f"rank{rank}_document"] = hit.get("document_title", "")
            result[f"rank{rank}_chunk_id"] = hit.get("chunk_id", "")
            result[f"rank{rank}_distance"] = hit.get("distance", "")
        result_rows.append(result)

        element_rows.append(
            {
                "eval_id": gold["eval_id"],
                "normalized_category": gold["normalized_category"],
                "difficulty": gold["difficulty"],
                "expected_core_elements": gold["expected_core_elements_raw"],
                "element_coverage_at_1": coverage_1,
                "matched_terms_at_1": "; ".join(matched_1),
                "match_reasons_at_1": "; ".join(reasons_1),
                "element_coverage_at_3": coverage_3,
                "matched_terms_at_3": "; ".join(matched_3),
                "match_reasons_at_3": "; ".join(reasons_3),
                "element_coverage_at_5": coverage_5,
                "matched_terms_at_5": "; ".join(matched_5),
                "match_reasons_at_5": "; ".join(reasons_5),
            }
        )

        for hit in hits:
            match = hit["document_match"]
            detail_rows.append(
                {
                    "eval_id": gold["eval_id"],
                    "source_question_id": gold["source_question_id"],
                    "normalized_category": gold["normalized_category"],
                    "difficulty": gold["difficulty"],
                    "question": gold["question"],
                    "expected_document_raw": gold["expected_document_raw"],
                    "expected_core_elements_raw": gold["expected_core_elements_raw"],
                    "rank": hit["rank"],
                    "document_title": hit["document_title"],
                    "chunk_id": hit["chunk_id"],
                    "source": hit["source"],
                    "page": hit["page"],
                    "distance": hit["distance"],
                    "document_text_excerpt": hit["document_text_excerpt"],
                    "metadata_json_safe": hit["metadata_json_safe"],
                    "automatic_document_match": "Y" if match["matched"] else "N",
                    "exact_normalized_match": match["exact_normalized_match"],
                    "title_token_match": match["title_token_match"],
                    "matched_expected_pattern": match["matched_expected_pattern"],
                    "match_reason": match["match_reason"],
                    "automatic_element_match": "; ".join(hit["matched_core_terms"]),
                    "core_match_reasons": "; ".join(hit["core_match_reasons"]),
                }
            )

    return result_rows, detail_rows, element_rows, failure_ids


def comparison_rows(
    baseline_metrics: dict[str, float | int],
    postfix_metrics: dict[str, float | int],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    all_rows = []
    rate_rows = []
    count_rows = []
    for key, label in {**RATE_METRIC_LABELS, **COUNT_METRIC_LABELS}.items():
        row = {
            "metric_key": key,
            "지표": label,
            "수정 전 자동 기준선": baseline_metrics[key],
            "수정 후 회귀평가": postfix_metrics[key],
            "증감": float(postfix_metrics[key]) - float(baseline_metrics[key]),
        }
        all_rows.append(row)
        compact = {
            "metric": label,
            "baseline": baseline_metrics[key],
            "postfix": postfix_metrics[key],
        }
        if key in RATE_METRIC_LABELS:
            rate_rows.append(compact)
        else:
            count_rows.append(compact)
    return all_rows, rate_rows, count_rows


def group_comparison(
    baseline_rows: list[dict[str, Any]],
    postfix_rows: list[dict[str, Any]],
    group_field: str,
) -> list[dict[str, Any]]:
    baseline = {row[group_field]: row for row in baseline_rows}
    postfix = {row[group_field]: row for row in postfix_rows}
    output = []
    for group in sorted(set(baseline) | set(postfix)):
        baseline_hit3 = float(baseline.get(group, {}).get("document_hit_at_3", 0))
        postfix_hit3 = float(postfix.get(group, {}).get("document_hit_at_3", 0))
        output.append(
            {
                group_field: group,
                "수정 전 Hit@3": baseline_hit3,
                "수정 후 Hit@3": postfix_hit3,
                "Hit@3 증감": postfix_hit3 - baseline_hit3,
                "수정 전 MRR": float(baseline.get(group, {}).get("mrr", 0)),
                "수정 후 MRR": float(postfix.get(group, {}).get("mrr", 0)),
            }
        )
    return output


def artifact_runtime_paths() -> tuple[Path, Path]:
    runtime_root = (
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
    )
    node_executable = runtime_root / "node" / "bin" / "node.exe"
    node_modules = runtime_root / "node" / "node_modules"
    if not node_executable.is_file() or not node_modules.is_dir():
        raise RuntimeError("Codex bundled artifact-tool runtime을 찾을 수 없습니다.")
    return node_executable, node_modules


def strip_xml_invalid_characters(value: str) -> str:
    """XLSX 직렬화 경계에서만 XML 1.0 금지 문자를 제거합니다."""
    return "".join(
        character
        for character in value
        if character in "\t\n\r"
        or "\x20" <= character <= "\ud7ff"
        or "\ue000" <= character <= "\ufffd"
        or "\U00010000" <= character <= "\U0010ffff"
    )


def sanitize_workbook_payload(value: Any) -> Any:
    if isinstance(value, str):
        cleaned = strip_xml_invalid_characters(value)
        if re.match(r"^[\t\r\n ]*[=+\-@]", cleaned):
            return f"'{cleaned}"
        return cleaned
    if isinstance(value, dict):
        return {key: sanitize_workbook_payload(item) for key, item in value.items()}
    if isinstance(value, list):
        return [sanitize_workbook_payload(item) for item in value]
    if isinstance(value, tuple):
        return [sanitize_workbook_payload(item) for item in value]
    return value


def build_workbooks(payload: dict[str, Any], qa_dir: Path | None) -> None:
    node_executable, node_modules = artifact_runtime_paths()
    with tempfile.TemporaryDirectory(prefix="minesafe_postfix_eval_xlsx_") as temporary:
        temporary_dir = Path(temporary)
        junction = temporary_dir / "node_modules"
        linked = subprocess.run(
            ["cmd", "/c", "mklink", "/J", str(junction), str(node_modules)],
            check=False,
            capture_output=True,
            text=True,
        )
        if linked.returncode != 0 or not junction.exists():
            raise RuntimeError("artifact-tool node_modules 임시 junction 생성 실패")
        payload_path = temporary_dir / "postfix_payload.json"
        temporary_builder = temporary_dir / BUILDER_PATH.name
        payload_path.write_text(
            json.dumps(sanitize_workbook_payload(payload), ensure_ascii=False),
            encoding="utf-8",
        )
        shutil.copyfile(BUILDER_PATH, temporary_builder)
        command = [
            str(node_executable),
            str(temporary_builder),
            str(payload_path),
            str(OUTPUT_DIR),
            str(qa_dir) if qa_dir else "",
        ]
        completed = subprocess.run(
            command,
            cwd=temporary_dir,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        if completed.returncode != 0:
            raise RuntimeError(
                "artifact-tool workbook 생성 또는 검증 실패: "
                + completed.stderr[-2000:]
            )


def format_metric_lines(
    metrics: dict[str, float | int],
    labels: dict[str, str],
) -> str:
    lines = []
    for key, label in labels.items():
        value = metrics[key]
        formatted = str(int(value)) if key in COUNT_METRIC_LABELS else f"{float(value):.4f}"
        lines.append(f"- {label}: {formatted}")
    return "\n".join(lines)


def format_change_lines(
    baseline: dict[str, float | int],
    postfix: dict[str, float | int],
) -> str:
    lines = []
    for key, label in {**RATE_METRIC_LABELS, **COUNT_METRIC_LABELS}.items():
        delta = float(postfix[key]) - float(baseline[key])
        formatted = f"{delta:+.0f}" if key in COUNT_METRIC_LABELS else f"{delta:+.4f}"
        lines.append(f"- {label}: {formatted}")
    return "\n".join(lines)


def format_group_lines(rows: list[dict[str, Any]], group_field: str) -> str:
    return "\n".join(
        f"- {row[group_field]}: 기준선 {row['수정 전 Hit@3']:.4f} / "
        f"수정 후 {row['수정 후 Hit@3']:.4f} / 증감 {row['Hit@3 증감']:+.4f}"
        for row in rows
    )


def write_quality_report(
    *,
    created_at: str,
    baseline_metrics: dict[str, float | int],
    postfix_metrics: dict[str, float | int],
    category_changes: list[dict[str, Any]],
    difficulty_changes: list[dict[str, Any]],
    review_ids: list[str],
    failure_ids: list[str],
    app_before: str,
    app_after: str,
    source_before: str,
    source_after: str,
    db_before: dict[str, Any],
    db_after: dict[str, Any],
) -> None:
    report = f"""MineSafe AI 수정 후 RAG 검색 회귀평가 품질 보고서

평가명: {EVALUATION_NAME}
평가 유형: {EVALUATION_TYPE}
생성 시각: {created_at}

[중요 주의]
{NOT_FINAL_LOCK_NOTICE}
{LEGAL_ACCURACY_NOTICE}

1. 평가 목적
- 기존 26_rag_retrieval_evaluation 결과는 수정 전 자동 기준선으로 보존한다.
- 현재 app.py 검색 로직으로 동일한 30문항을 다시 실행해 변화만 확인한다.
- 질문 내용·ID·유형·난이도·기대 문서·기대 요소·Top K·지표 계산식을 변경하지 않았다.

2. 실행 환경
- Python: {sys.executable}
- embedding 모델: {EMBEDDING_MODEL_NAME}
- embedding 차원: {EMBEDDING_DIMENSION}
- normalize_embeddings: {NORMALIZE_EMBEDDINGS}
- Vector DB: 10_vector_db_with_major_accident_docs (평가 목적의 로컬 읽기 질의만 수행)
- collection: {COLLECTION_NAME}
- Top K: {TOP_K}
- 외부 Gemini·Naver·인터넷 호출: 없음

3. 버전 무결성
- 수정 전 기준선 app SHA-256: {BASELINE_APP_SHA256}
- 수정 후 app SHA-256 작업 전/후: {app_before} / {app_after}
- 원본 110문항 SHA-256 작업 전/후: {source_before} / {source_after}
- Gold set SHA-256: {GOLD_SET_SHA256}
- DB collection count 작업 전/후: {db_before['count']} / {db_after['count']}
- DB ID SHA-256 작업 전/후: {db_before['ids_sha256']} / {db_after['ids_sha256']}
- 물리 파일 주의: ChromaDB 읽기 질의 과정에서 data_level0.bin, length.bin, chroma.sqlite3의 수정시각이 평가 시각으로 갱신됐다. 논리 count와 ID SHA는 동일하지만 작업 전 물리 파일 SHA를 수집하지 않아 물리 바이트 무변경은 단정하지 않는다.

4. 기존 기준선 지표
{format_metric_lines(baseline_metrics, {**RATE_METRIC_LABELS, **COUNT_METRIC_LABELS})}

5. 수정 후 회귀평가 지표
{format_metric_lines(postfix_metrics, {**RATE_METRIC_LABELS, **COUNT_METRIC_LABELS})}

6. 기존 대비 증감
{format_change_lines(baseline_metrics, postfix_metrics)}

7. 유형별 Hit@3 변화
{format_group_lines(category_changes, 'category')}

8. 난이도별 Hit@3 변화
{format_group_lines(difficulty_changes, 'difficulty')}

9. 수동 검토 필요 문항
- {len(review_ids)}문항: {', '.join(review_ids) if review_ids else '없음'}

10. 검색 실행 실패 문항
- {len(failure_ids)}문항: {', '.join(failure_ids) if failure_ids else '없음'}

11. 산출물
- rag_retrieval_postfix_results_30.tsv
- rag_retrieval_postfix_results_30.xlsx
- rag_retrieval_postfix_manual_review_30.xlsx
- rag_retrieval_postfix_summary.xlsx
- rag_retrieval_baseline_vs_postfix.xlsx
- rag_retrieval_baseline_vs_postfix.png
- rag_retrieval_postfix_quality_report.txt
- evaluation_manifest.json

12. 평가 한계
- 동일한 30문항을 재사용했으므로 최종 잠금 성능 평가나 독립·블라인드 평가가 아니다.
- Hit@K와 MRR은 기대 문서 검색 순위, Element Coverage는 규칙 기반 용어 출현을 측정한다.
- 검색 결과가 좋아져도 법령 해석, 현장 사실관계, 최종 답변의 법적·실무적 정확성을 자동으로 보장하지 않는다.
- 자동 REVIEW 판정 문항은 수동 검토 workbook에서 사람이 검색 chunk를 확인해야 한다.
"""
    QUALITY_REPORT_PATH.write_text(report, encoding="utf-8")


def build_manifest(
    *,
    created_at: str,
    baseline_hashes: dict[str, str],
    baseline_metrics: dict[str, float | int],
    postfix_metrics: dict[str, float | int],
    category_changes: list[dict[str, Any]],
    difficulty_changes: list[dict[str, Any]],
    question_ids: list[str],
    review_ids: list[str],
    failure_ids: list[str],
    db_snapshot: dict[str, Any],
) -> dict[str, Any]:
    result_hashes = {
        path.name: sha256_file(path)
        for path in sorted(OUTPUT_DIR.iterdir(), key=lambda item: item.name)
        if path.is_file() and path != MANIFEST_PATH
    }
    metric_deltas = {
        key: float(postfix_metrics[key]) - float(baseline_metrics[key])
        for key in {**RATE_METRIC_LABELS, **COUNT_METRIC_LABELS}
    }
    return {
        "evaluation_name": EVALUATION_NAME,
        "evaluation_type": EVALUATION_TYPE,
        "created_at": created_at,
        "baseline_folder": "26_rag_retrieval_evaluation",
        "postfix_folder": "27_rag_retrieval_evaluation_postfix",
        "baseline_app_sha256": BASELINE_APP_SHA256,
        "postfix_app_sha256": POSTFIX_APP_SHA256,
        "original_question_set_sha256": ORIGINAL_QUESTION_SET_SHA256,
        "gold_set_sha256": GOLD_SET_SHA256,
        "evaluation_question_count": QUESTION_COUNT,
        "evaluation_question_ids": question_ids,
        "evaluation_script_sha256": sha256_file(Path(__file__)),
        "workbook_builder_sha256": sha256_file(BUILDER_PATH),
        "baseline_files_sha256": baseline_hashes,
        "postfix_result_files_sha256": result_hashes,
        "python_executable": sys.executable,
        "embedding_model": EMBEDDING_MODEL_NAME,
        "embedding_dimension": EMBEDDING_DIMENSION,
        "normalize_embeddings": NORMALIZE_EMBEDDINGS,
        "vector_db": "10_vector_db_with_major_accident_docs",
        "collection_name": COLLECTION_NAME,
        "collection_count": db_snapshot["count"],
        "top_k": TOP_K,
        "baseline_metrics": baseline_metrics,
        "postfix_metrics": postfix_metrics,
        "metric_deltas": metric_deltas,
        "category_hit_at_3_changes": category_changes,
        "difficulty_hit_at_3_changes": difficulty_changes,
        "manual_review_count": len(review_ids),
        "manual_review_ids": review_ids,
        "search_failure_count": len(failure_ids),
        "search_failure_ids": failure_ids,
        "same_30_questions_reused": True,
        "external_api_or_internet_used": False,
        "vector_db_logical_integrity": "평가 전후 collection count와 ID SHA-256 일치",
        "physical_query_side_effect_notice": (
            "ChromaDB 읽기 질의 과정에서 data_level0.bin, length.bin, chroma.sqlite3의 "
            "수정시각이 평가 시각으로 갱신됐다. 작업 전 물리 SHA가 없어 물리 바이트 "
            "무변경은 단정하지 않는다."
        ),
        "not_final_lock_notice": NOT_FINAL_LOCK_NOTICE,
        "legal_accuracy_notice": LEGAL_ACCURACY_NOTICE,
    }


def run_evaluation(qa_dir: Path | None = None) -> dict[str, Any]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    baseline_hashes_before = assert_baseline_files_unchanged()
    app_before = sha256_file(APP_PATH)
    source_before = sha256_file(SOURCE_QUESTIONS_PATH)
    gold_sha = sha256_file(GOLD_PATH)
    if app_before != POSTFIX_APP_SHA256:
        raise RuntimeError(f"현재 app.py SHA-256 불일치: {app_before}")
    if source_before != ORIGINAL_QUESTION_SET_SHA256:
        raise RuntimeError(f"원본 110문항 SHA-256 불일치: {source_before}")
    if gold_sha != GOLD_SET_SHA256:
        raise RuntimeError(f"기존 Gold set SHA-256 불일치: {gold_sha}")

    base = load_baseline_module()
    base.OPERATIONAL_FUNCTIONS = set(base.OPERATIONAL_FUNCTIONS) | {
        "extract_query_context_signals"
    }
    gold_rows = read_tsv(GOLD_PATH)
    expected_ids = [f"R{index:03d}" for index in range(1, QUESTION_COUNT + 1)]
    if len(gold_rows) != QUESTION_COUNT or [row["eval_id"] for row in gold_rows] != expected_ids:
        raise RuntimeError("기존 30문항 Gold set의 문항 수 또는 순서가 다릅니다.")

    baseline_result_rows = read_tsv(BASELINE_RESULT_PATH)
    baseline_metrics = metrics_from_rows(baseline_result_rows)
    validate_baseline_metrics(baseline_metrics)

    namespace = base.load_operational_search_namespace(APP_PATH)
    base.validate_operational_contract(namespace)

    import chromadb

    client = chromadb.PersistentClient(path=str(DB_PATH))
    collection = client.get_collection(name=COLLECTION_NAME)
    db_before = base.snapshot_collection_logical(collection)
    if db_before["count"] != 820 or db_before["id_count"] != 820:
        raise RuntimeError(f"collection count가 820이 아닙니다: {db_before['count']}")
    model = base.load_local_embedding_model(EMBEDDING_MODEL_NAME)

    result_rows, detail_rows, element_rows, failure_ids = evaluate_questions(
        base,
        gold_rows,
        collection,
        model,
        namespace,
    )
    postfix_metrics = metrics_from_rows(result_rows)
    review_ids = [
        row["eval_id"]
        for row in result_rows
        if row["auto_review_status"] == "REVIEW"
    ]

    base.write_tsv(RESULT_TSV_PATH, result_rows, base.RESULT_COLUMNS)
    postfix_category_rows_raw = base.aggregate_results(result_rows, "category")
    postfix_category_rows = [
        {
            "normalized_category": row["category"],
            **{key: value for key, value in row.items() if key != "category"},
        }
        for row in postfix_category_rows_raw
    ]
    postfix_difficulty_rows = base.aggregate_results(result_rows, "difficulty")
    baseline_category_rows = base.aggregate_results(baseline_result_rows, "category")
    baseline_difficulty_rows = base.aggregate_results(baseline_result_rows, "difficulty")
    category_changes = group_comparison(
        baseline_category_rows,
        postfix_category_rows_raw,
        "category",
    )
    difficulty_changes = group_comparison(
        baseline_difficulty_rows,
        postfix_difficulty_rows,
        "difficulty",
    )
    frequency_counter = Counter(
        row["document_title"]
        for row in detail_rows
        if row["document_title"]
    )
    frequency_rows = [
        {"document_title": title, "retrieval_count": count}
        for title, count in frequency_counter.most_common()
    ]
    created_at = datetime.now().astimezone().isoformat(timespec="seconds")
    summary_rows = [
        {"항목": "평가명", "값": EVALUATION_NAME},
        {"항목": "평가 유형", "값": EVALUATION_TYPE},
        {"항목": "평가 문항 수", "값": QUESTION_COUNT},
        {"항목": "동일 30문항 재사용", "값": "예"},
        {"항목": "최종 잠금 평가 여부", "값": "아님"},
        {"항목": "DB collection count", "값": db_before["count"]},
        {"항목": "embedding 모델", "값": EMBEDDING_MODEL_NAME},
        {"항목": "Document Hit@1", "값": postfix_metrics["document_hit_at_1"]},
        {"항목": "Document Hit@3", "값": postfix_metrics["document_hit_at_3"]},
        {"항목": "Document Hit@5", "값": postfix_metrics["document_hit_at_5"]},
        {"항목": "MRR", "값": postfix_metrics["mrr"]},
        {"항목": "Element Coverage@1", "값": postfix_metrics["element_coverage_at_1"]},
        {"항목": "Element Coverage@3", "값": postfix_metrics["element_coverage_at_3"]},
        {"항목": "Element Coverage@5", "값": postfix_metrics["element_coverage_at_5"]},
        {"항목": "Metadata 완전성", "값": postfix_metrics["metadata_completeness"]},
        {"항목": "수동 검토 필요 문항 수", "값": len(review_ids)},
        {"항목": "검색 실행 실패 문항 수", "값": len(failure_ids)},
        {"항목": "수정 후 app SHA-256", "값": POSTFIX_APP_SHA256},
        {"항목": "Gold set SHA-256", "값": GOLD_SET_SHA256},
        {
            "항목": "평가 실행 시각",
            "값": "KST " + created_at.replace("T", " ").replace("+09:00", ""),
        },
    ]

    payload = base.make_workbook_payload(
        gold_rows,
        result_rows,
        detail_rows,
        postfix_category_rows,
        postfix_difficulty_rows,
        frequency_rows,
        element_rows,
        summary_rows,
    )
    payload["methodRows"] = [
        {"구분": "기존 결과", "내용": "26번 폴더 결과는 수정 전 자동 기준선이며 변경하지 않았다."},
        {"구분": "신규 결과", "내용": "질문 context 및 미완성 답변 안전장치 수정 후 회귀평가다."},
        {"구분": "질문 재사용", "내용": "동일한 30문항·질문 ID·유형·난이도·기대 문서·기대 요소를 그대로 재사용했다."},
        {"구분": "평가 제한", "내용": NOT_FINAL_LOCK_NOTICE},
        {"구분": "지표 해석", "내용": LEGAL_ACCURACY_NOTICE},
        {"구분": "외부 호출", "내용": "Gemini API, Naver API 및 인터넷을 호출하지 않았다."},
    ]
    payload["interpretationRows"] = [
        {"지표": "Document Hit@K", "해석": "기대 문서가 상위 K개 검색 결과에 포함된 비율"},
        {"지표": "MRR", "해석": "기대 문서가 검색 결과의 앞 순위에 배치되는 정도"},
        {"지표": "Element Coverage@K", "해석": "기대 안전조치 용어가 상위 K개 chunk에 포함된 규칙 기반 보조 지표"},
        {"지표": "Metadata 완전성", "해석": "문서명·chunk_id·source·page 정보가 존재하는 비율"},
        {"지표": "주의", "해석": NOT_FINAL_LOCK_NOTICE + " " + LEGAL_ACCURACY_NOTICE},
    ]
    all_comparison_rows, rate_comparison_rows, count_comparison_rows = comparison_rows(
        baseline_metrics,
        postfix_metrics,
    )
    payload.update(
        {
            "comparisonHeaders": ["metric_key", "지표", "수정 전 자동 기준선", "수정 후 회귀평가", "증감"],
            "comparisonRows": all_comparison_rows,
            "comparisonRateRows": rate_comparison_rows,
            "comparisonCountRows": count_comparison_rows,
            "categoryComparisonHeaders": ["category", "수정 전 Hit@3", "수정 후 Hit@3", "Hit@3 증감", "수정 전 MRR", "수정 후 MRR"],
            "categoryComparisonRows": category_changes,
            "difficultyComparisonHeaders": ["difficulty", "수정 전 Hit@3", "수정 후 Hit@3", "Hit@3 증감", "수정 전 MRR", "수정 후 MRR"],
            "difficultyComparisonRows": difficulty_changes,
        }
    )
    build_workbooks(payload, qa_dir)
    for sidecar_path in OUTPUT_DIR.glob("*.xlsx.inspect.ndjson"):
        sidecar_path.unlink()

    db_after = base.snapshot_collection_logical(collection)
    del model, collection, client
    gc.collect()
    app_after = sha256_file(APP_PATH)
    source_after = sha256_file(SOURCE_QUESTIONS_PATH)
    baseline_hashes_after = assert_baseline_files_unchanged()
    if app_after != app_before:
        raise RuntimeError("평가 전후 app.py SHA-256이 달라졌습니다.")
    if source_after != source_before:
        raise RuntimeError("평가 전후 원본 질문 세트 SHA-256이 달라졌습니다.")
    if baseline_hashes_after != baseline_hashes_before:
        raise RuntimeError("평가 전후 기존 26번 기준선 파일이 달라졌습니다.")
    if (
        db_after["count"] != db_before["count"]
        or db_after["ids_sha256"] != db_before["ids_sha256"]
    ):
        raise RuntimeError("평가 전후 Vector DB 논리 데이터가 달라졌습니다.")

    write_quality_report(
        created_at=created_at,
        baseline_metrics=baseline_metrics,
        postfix_metrics=postfix_metrics,
        category_changes=category_changes,
        difficulty_changes=difficulty_changes,
        review_ids=review_ids,
        failure_ids=failure_ids,
        app_before=app_before,
        app_after=app_after,
        source_before=source_before,
        source_after=source_after,
        db_before=db_before,
        db_after=db_after,
    )
    manifest = build_manifest(
        created_at=created_at,
        baseline_hashes=baseline_hashes_after,
        baseline_metrics=baseline_metrics,
        postfix_metrics=postfix_metrics,
        category_changes=category_changes,
        difficulty_changes=difficulty_changes,
        question_ids=[row["eval_id"] for row in gold_rows],
        review_ids=review_ids,
        failure_ids=failure_ids,
        db_snapshot=db_after,
    )
    MANIFEST_PATH.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return manifest


def refresh_manifest_hashes_only() -> dict[str, Any]:
    """Vector DB를 다시 열지 않고 현재 산출물과 실행 스크립트 해시만 갱신합니다."""
    if not MANIFEST_PATH.is_file():
        raise RuntimeError("갱신할 evaluation_manifest.json이 없습니다.")
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    manifest["evaluation_script_sha256"] = sha256_file(Path(__file__))
    manifest["workbook_builder_sha256"] = sha256_file(BUILDER_PATH)
    manifest["postfix_result_files_sha256"] = {
        path.name: sha256_file(path)
        for path in sorted(OUTPUT_DIR.iterdir(), key=lambda item: item.name)
        if path.is_file() and path != MANIFEST_PATH
    }
    manifest["vector_db_logical_integrity"] = (
        "평가 전후 collection count와 ID SHA-256 일치"
    )
    manifest["physical_query_side_effect_notice"] = (
        "ChromaDB 읽기 질의 과정에서 data_level0.bin, length.bin, chroma.sqlite3의 "
        "수정시각이 평가 시각으로 갱신됐다. 작업 전 물리 SHA가 없어 물리 바이트 "
        "무변경은 단정하지 않는다."
    )
    MANIFEST_PATH.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return manifest


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="MineSafe AI 동일 30문항 수정 후 RAG 검색 회귀평가"
    )
    parser.add_argument(
        "--qa-dir",
        type=Path,
        default=None,
        help="artifact-tool 전체 시트 렌더와 검사 결과를 저장할 프로젝트 밖 QA 폴더",
    )
    parser.add_argument(
        "--refresh-manifest-hashes-only",
        action="store_true",
        help="Vector DB를 열지 않고 현재 스크립트·산출물 해시와 물리 부작용 주의문만 갱신",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest = (
        refresh_manifest_hashes_only()
        if args.refresh_manifest_hashes_only
        else run_evaluation(args.qa_dir)
    )
    print(
        json.dumps(
            {
                "evaluation_name": manifest["evaluation_name"],
                "evaluation_type": manifest["evaluation_type"],
                "postfix_metrics": manifest["postfix_metrics"],
                "metric_deltas": manifest["metric_deltas"],
                "manual_review_count": manifest["manual_review_count"],
                "search_failure_count": manifest["search_failure_count"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
