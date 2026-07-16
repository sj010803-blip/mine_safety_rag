from __future__ import annotations

import argparse
import ast
import csv
from collections import Counter, defaultdict
from datetime import datetime
import gc
import hashlib
import importlib.metadata
import io
import json
from pathlib import Path
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
from typing import Any


ROOT_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT_DIR / "26_rag_retrieval_evaluation"
SOURCE_QUESTIONS_PATH = ROOT_DIR / "02_질문시나리오" / "question_scenarios_110.tsv"
APP_PATH = ROOT_DIR / "app.py"
DB_PATH = ROOT_DIR / "10_vector_db_with_major_accident_docs"
COLLECTION_NAME = "mine_safety_docs"
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
EMBEDDING_DIMENSION = 384
NORMALIZE_EMBEDDINGS = True
TOP_K = 5
RANDOM_SEED = 20260716
GOLD_SET_VERSION = "RAG-RETRIEVAL-30-V1"
EXPECTED_SOURCE_SHA256 = "544888A7717DD31CB0AF5099128D8493DE7AD50EB38F9ACE90FEFA2AFD72277A"
EXPECTED_APP_SHA256 = "983B80A2C390829BA7750830696BC333E8AEABC1A7A4F7407237995AB4AC2FF2"

GOLD_PATH = OUTPUT_DIR / "rag_retrieval_eval_questions_30.tsv"
RESULT_TSV_PATH = OUTPUT_DIR / "rag_retrieval_eval_results_30.tsv"
RESULT_XLSX_PATH = OUTPUT_DIR / "rag_retrieval_eval_results_30.xlsx"
MANUAL_XLSX_PATH = OUTPUT_DIR / "rag_retrieval_manual_review_30.xlsx"
SUMMARY_XLSX_PATH = OUTPUT_DIR / "rag_retrieval_summary.xlsx"
HIT_CHART_PATH = OUTPUT_DIR / "rag_retrieval_hit_at_k.png"
CATEGORY_CHART_PATH = OUTPUT_DIR / "rag_retrieval_category_hit3.png"
DIFFICULTY_CHART_PATH = OUTPUT_DIR / "rag_retrieval_difficulty_hit3.png"
REPORT_PATH = OUTPUT_DIR / "rag_retrieval_quality_report.txt"
TEST_PATH = ROOT_DIR / "18_legal_evidence_features" / "test_rag_retrieval_quality_evaluation.py"

REQUIRED_SOURCE_COLUMNS = {
    "번호",
    "분류",
    "난이도",
    "질문 시나리오",
    "기대 검색 문서",
    "정답에 포함되어야 할 핵심 요소",
}

SELECTION_PLAN: tuple[tuple[str, tuple[int, int, int]], ...] = (
    ("발파·불발", (3, 4, 35)),
    ("환기·메탄·유해가스·산소부족", (5, 7, 8)),
    ("분진·호흡보호·보호구", (13, 14, 28)),
    ("낙반·붕락·지보", (2, 40, 54)),
    ("전기안전·에너지 차단", (12, 66, 68)),
    ("장비·운반·차량", (9, 10, 84)),
    ("컨베이어·끼임", (11, 83, 89)),
    ("사고보고·응급조치", (20, 76, 78)),
    ("위험성평가·KRAS·TBM", (24, 31, 73)),
    ("중대재해처벌법·증빙자료", (21, 80, 93)),
)

GOLD_COLUMNS = [
    "eval_id",
    "source_question_id",
    "category",
    "normalized_category",
    "difficulty",
    "question",
    "expected_document_raw",
    "expected_document_patterns",
    "expected_core_elements_raw",
    "expected_core_terms",
    "selection_reason",
    "source_file",
    "gold_set_version",
]

RESULT_COLUMNS = [
    "eval_id",
    "source_question_id",
    "category",
    "difficulty",
    "question",
    "expected_document_raw",
    "expected_document_patterns",
    "expected_core_elements_raw",
    "expected_core_terms",
    "gold_set_sha256",
    "rank1_document",
    "rank1_chunk_id",
    "rank1_distance",
    "rank2_document",
    "rank2_chunk_id",
    "rank2_distance",
    "rank3_document",
    "rank3_chunk_id",
    "rank3_distance",
    "rank4_document",
    "rank4_chunk_id",
    "rank4_distance",
    "rank5_document",
    "rank5_chunk_id",
    "rank5_distance",
    "first_expected_document_rank",
    "document_hit_at_1",
    "document_hit_at_3",
    "document_hit_at_5",
    "reciprocal_rank",
    "element_coverage_at_1",
    "element_coverage_at_3",
    "element_coverage_at_5",
    "unique_document_count_top5",
    "metadata_completeness_top5",
    "auto_review_status",
    "auto_review_reason",
]

DOCUMENT_ALIASES = {
    "광산안전지침": ("광산안전업무처리지침",),
    "위험성평가안내서": ("새로운위험성평가안내서",),
    "중대재해처벌법": ("중대재해처벌등에관한법률",),
}

CORE_TERM_ALIASES = {
    "작업중지": ("작업중지", "작업 중지", "작업을 중단", "운전 정지"),
    "대피": ("대피", "피난", "작업자 철수"),
    "접근통제": ("접근통제", "접근 통제", "출입통제", "출입 통제", "접근금지"),
    "전원차단": ("전원차단", "전원 차단", "전기 차단", "전로 차단"),
    "잠금표지": ("잠금표지", "잠금·표지", "잠금 및 표지", "잠금조치", "LOTO"),
    "환기": ("환기", "통기", "송풍", "배기"),
    "재측정": ("재측정", "다시 측정", "측정값 재확인"),
    "기록관리": ("기록관리", "기록 관리", "기록", "보관"),
    "안전관리자보고": ("안전관리자 보고", "관리자 보고", "책임자 보고", "즉시 보고"),
    "위험성평가": ("위험성평가", "위험성 평가", "위험요인 평가"),
    "개선조치": ("개선조치", "개선 조치", "시정조치"),
    "작업재개확인": ("작업재개", "작업 재개", "재개 승인", "재개 전 확인"),
}

OPERATIONAL_FUNCTIONS = {
    "get_source",
    "get_chunk_id",
    "clean_text",
    "classify_question_type",
    "detect_question_intent",
    "detect_ppe_item",
    "expand_search_query",
    "squared_l2_distance",
    "make_search_result",
    "source_matches_markers",
    "add_type_source_candidates",
    "rerank_search_results",
}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest().upper()


def directory_metadata_snapshot(path: Path) -> dict[str, Any]:
    records: list[str] = []
    total_bytes = 0
    for file_path in sorted(candidate for candidate in path.rglob("*") if candidate.is_file()):
        stat = file_path.stat()
        total_bytes += stat.st_size
        records.append(
            f"{file_path.relative_to(path).as_posix()}\t{stat.st_size}\t{stat.st_mtime_ns}"
        )
    return {
        "file_count": len(records),
        "total_bytes": total_bytes,
        "metadata_sha256": sha256_text("\n".join(records)),
    }


def resolve_scenario_columns(fieldnames: list[str] | None) -> dict[str, str]:
    if not fieldnames:
        raise RuntimeError("110문항 원본 헤더를 읽을 수 없습니다.")
    normalized = {re.sub(r"\s+", "", name): name for name in fieldnames}
    aliases = {
        "번호": ("번호", "문항번호", "id"),
        "분류": ("분류", "유형", "카테고리"),
        "난이도": ("난이도", "수준"),
        "질문 시나리오": ("질문시나리오", "질문", "시나리오"),
        "기대 검색 문서": ("기대검색문서", "기대문서", "정답문서"),
        "정답에 포함되어야 할 핵심 요소": (
            "정답에포함되어야할핵심요소",
            "핵심요소",
            "기대핵심요소",
        ),
    }
    resolved: dict[str, str] = {}
    for canonical, candidates in aliases.items():
        for candidate in candidates:
            key = re.sub(r"\s+", "", candidate)
            if key in normalized:
                resolved[canonical] = normalized[key]
                break
        if canonical not in resolved:
            raise RuntimeError(f"필수 열 누락: {canonical}")
    return resolved


def read_scenario_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        columns = resolve_scenario_columns(reader.fieldnames)
        rows = []
        for source in reader:
            rows.append({canonical: str(source.get(actual, "")).strip() for canonical, actual in columns.items()})
    if len(rows) != 110:
        raise RuntimeError(f"원본 문항 수가 110이 아닙니다: {len(rows)}")
    return rows


def normalize_difficulty(value: str) -> str:
    compact = re.sub(r"\s+", "", value)
    mapping = {
        "하": "하",
        "낮음": "하",
        "초급": "하",
        "중": "중",
        "보통": "중",
        "중급": "중",
        "상": "상",
        "높음": "상",
        "고급": "상",
    }
    if compact not in mapping:
        raise RuntimeError(f"알 수 없는 난이도: {value}")
    return mapping[compact]


def split_expected_documents(raw: str) -> list[str]:
    prepared = re.sub(r"(?:\r?\n|;|；|,|，|/|·|ㆍ)", "|", raw)
    prepared = re.sub(r"\s+(?:또는|및)\s+", "|", prepared)
    return [part.strip(" -•\t") for part in prepared.split("|") if part.strip(" -•\t")]


def normalize_document_title(value: str) -> str:
    text = Path(str(value).strip()).name
    text = re.sub(r"\.(?:txt|pdf|docx?|hwp|xlsx?)$", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^\d{1,3}[_\-\s]+", "", text)
    text = re.sub(r"[\[\](){}]", " ", text)
    text = re.sub(r"[^0-9A-Za-z가-힣]+", "", text)
    return text.lower()


def document_title_tokens(value: str) -> list[str]:
    text = Path(str(value).strip()).stem
    text = re.sub(r"^\d{1,3}[_\-\s]+", "", text)
    tokens = [token.lower() for token in re.split(r"[^0-9A-Za-z가-힣]+", text) if len(token) >= 2]
    return list(dict.fromkeys(tokens))


def build_expected_document_patterns(raw: str) -> list[dict[str, Any]]:
    patterns = []
    for title in split_expected_documents(raw):
        normalized = normalize_document_title(title)
        aliases = list(DOCUMENT_ALIASES.get(normalized, ()))
        patterns.append(
            {
                "raw": title,
                "normalized": normalized,
                "aliases": [normalize_document_title(alias) for alias in aliases],
                "tokens": document_title_tokens(title),
            }
        )
    return patterns


def split_expected_core_elements(raw: str) -> list[str]:
    prepared = re.sub(r"(?:\r?\n|;|；|,|，|/|·|ㆍ)", "|", raw)
    prepared = re.sub(r"(?:^|\s)\d+[.)]\s*", "|", prepared)
    prepared = re.sub(r"\s+(?:또는|및)\s+", "|", prepared)
    return [part.strip(" -•\t") for part in prepared.split("|") if part.strip(" -•\t")]


def normalize_term(value: str) -> str:
    return re.sub(r"[^0-9A-Za-z가-힣]+", "", value).lower()


def build_expected_core_terms(raw: str) -> list[dict[str, Any]]:
    terms = []
    for term in split_expected_core_elements(raw):
        normalized = normalize_term(term)
        aliases = [term]
        if normalized in CORE_TERM_ALIASES:
            aliases.extend(CORE_TERM_ALIASES[normalized])
        terms.append(
            {
                "term": term,
                "normalized": normalized,
                "aliases": list(dict.fromkeys(alias for alias in aliases if alias)),
            }
        )
    return terms


def match_expected_document(retrieved_title: str, patterns: list[dict[str, Any]]) -> dict[str, Any]:
    retrieved_normalized = normalize_document_title(retrieved_title)
    retrieved_tokens = set(document_title_tokens(retrieved_title))
    for pattern in patterns:
        candidates = [pattern.get("normalized", ""), *pattern.get("aliases", [])]
        for candidate in candidates:
            if candidate and (
                candidate == retrieved_normalized
                or (len(candidate) >= 4 and candidate in retrieved_normalized)
                or (len(retrieved_normalized) >= 4 and retrieved_normalized in candidate)
            ):
                return {
                    "matched": True,
                    "exact_normalized_match": True,
                    "title_token_match": False,
                    "matched_expected_pattern": pattern.get("raw", ""),
                    "match_reason": "정규화 제목 포함 일치",
                }
        expected_tokens = set(pattern.get("tokens", []))
        if len(expected_tokens) >= 2:
            ratio = len(expected_tokens & retrieved_tokens) / len(expected_tokens)
            if ratio >= 0.75:
                return {
                    "matched": True,
                    "exact_normalized_match": False,
                    "title_token_match": True,
                    "matched_expected_pattern": pattern.get("raw", ""),
                    "match_reason": f"핵심 제목 토큰 {ratio:.2f} 일치",
                }
    return {
        "matched": False,
        "exact_normalized_match": False,
        "title_token_match": False,
        "matched_expected_pattern": "",
        "match_reason": "기대 문서 제목 불일치",
    }


def match_core_term(text: str, term: dict[str, Any]) -> tuple[bool, str]:
    normalized_text = normalize_term(text)
    for alias in term.get("aliases", []):
        normalized_alias = normalize_term(alias)
        if normalized_alias and normalized_alias in normalized_text:
            return True, f"alias:{alias}"
    return False, ""


def serialize_tsv(rows: list[dict[str, Any]], fieldnames: list[str]) -> bytes:
    stream = io.StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=fieldnames, delimiter="\t", lineterminator="\n", extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return ("\ufeff" + stream.getvalue()).encode("utf-8")


def freeze_or_validate_gold_set(path: Path, content: bytes) -> str:
    if path.exists():
        existing = path.read_bytes()
        if existing != content:
            raise RuntimeError("동결된 30문항 gold set과 재생성 결과가 다릅니다. 덮어쓰지 않습니다.")
    else:
        path.write_bytes(content)
    return hashlib.sha256(content).hexdigest().upper()


def select_balanced_questions(rows: list[dict[str, str]], seed: int = RANDOM_SEED) -> list[dict[str, Any]]:
    if seed != RANDOM_SEED:
        raise RuntimeError(f"고정 seed는 {RANDOM_SEED}이어야 합니다.")
    indexed = {int(row["번호"]): row for row in rows}
    selected: list[dict[str, Any]] = []
    eval_index = 1
    for normalized_category, source_ids in SELECTION_PLAN:
        for source_id in source_ids:
            if source_id not in indexed:
                raise RuntimeError(f"선정 문항 누락: {source_id}")
            source = indexed[source_id]
            expected_document_raw = source["기대 검색 문서"]
            expected_core_raw = source["정답에 포함되어야 할 핵심 요소"]
            if not expected_document_raw or not expected_core_raw:
                raise RuntimeError(f"선정 문항의 기대값 누락: {source_id}")
            if normalized_category == "컨베이어·끼임" and source_id in {83, 89}:
                selection_reason = (
                    "컨베이어 직접 문항 부족으로 정비 전 에너지 차단·잠금표지 인접 문항을 "
                    f"고정 seed {seed} 계획으로 보충"
                )
            else:
                selection_reason = f"목표 유형 직접 문항을 고정 seed {seed} 계획으로 선정"
            selected.append(
                {
                    "eval_id": f"R{eval_index:03d}",
                    "source_question_id": str(source_id),
                    "category": source["분류"],
                    "normalized_category": normalized_category,
                    "difficulty": normalize_difficulty(source["난이도"]),
                    "question": source["질문 시나리오"],
                    "expected_document_raw": expected_document_raw,
                    "expected_document_patterns": json.dumps(
                        build_expected_document_patterns(expected_document_raw), ensure_ascii=False, separators=(",", ":")
                    ),
                    "expected_core_elements_raw": expected_core_raw,
                    "expected_core_terms": json.dumps(
                        build_expected_core_terms(expected_core_raw), ensure_ascii=False, separators=(",", ":")
                    ),
                    "selection_reason": selection_reason,
                    "source_file": "02_질문시나리오/question_scenarios_110.tsv",
                    "gold_set_version": GOLD_SET_VERSION,
                }
            )
            eval_index += 1
    if len(selected) != 30 or len({row["source_question_id"] for row in selected}) != 30:
        raise RuntimeError("30문항 선정 수 또는 원본 ID 유일성 오류")
    return selected


def load_operational_search_namespace(app_path: Path) -> dict[str, Any]:
    tree = ast.parse(app_path.read_text(encoding="utf-8-sig"), filename=str(app_path))
    definitions: dict[str, ast.AST] = {}
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            definitions[node.name] = node
        elif isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            for target in targets:
                if isinstance(target, ast.Name):
                    definitions[target.id] = node

    contract_constants = {"COLLECTION_NAME", "EMBEDDING_MODEL_NAME", "RERANK_TYPE_CONFIGS"}
    needed = set(OPERATIONAL_FUNCTIONS) | contract_constants
    queue = list(needed)
    while queue:
        name = queue.pop()
        node = definitions.get(name)
        if node is None:
            raise RuntimeError(f"app.py 운영 검색 계약 누락: {name}")
        for child in ast.walk(node):
            if isinstance(child, ast.Name) and isinstance(child.ctx, ast.Load):
                dependency = child.id
                if dependency in definitions and dependency not in needed:
                    needed.add(dependency)
                    queue.append(dependency)

    selected_nodes = sorted({id(definitions[name]): definitions[name] for name in needed}.values(), key=lambda item: item.lineno)
    module = ast.Module(body=selected_nodes, type_ignores=[])
    ast.fix_missing_locations(module)
    namespace: dict[str, Any] = {"Any": Any}
    exec(compile(module, str(app_path), "exec"), namespace)
    for name in OPERATIONAL_FUNCTIONS:
        if not callable(namespace.get(name)):
            raise RuntimeError(f"운영 검색 함수 로드 실패: {name}")
    return namespace


def validate_operational_contract(namespace: dict[str, Any]) -> None:
    expected = {
        "EMBEDDING_MODEL_NAME": EMBEDDING_MODEL_NAME,
        "COLLECTION_NAME": COLLECTION_NAME,
    }
    for name, value in expected.items():
        if namespace.get(name) != value:
            raise RuntimeError(f"app.py 운영 설정 불일치: {name}")


def snapshot_collection_logical(collection: Any) -> dict[str, Any]:
    rows = collection.get(include=["metadatas"])
    identifiers = sorted(str(value) for value in rows.get("ids", []))
    metadata_keys = sorted(
        {
            str(key)
            for metadata in (rows.get("metadatas") or [])
            if isinstance(metadata, dict)
            for key in metadata
        }
    )
    return {
        "count": int(collection.count()),
        "id_count": len(identifiers),
        "ids_sha256": sha256_text("\n".join(identifiers)),
        "metadata_fields": metadata_keys,
    }


def load_local_embedding_model(model_name: str) -> Any:
    from sentence_transformers import SentenceTransformer

    try:
        return SentenceTransformer(model_name, local_files_only=True)
    except Exception as error:
        raise RuntimeError("로컬 embedding 모델을 불러올 수 없어 평가를 중단합니다.") from error


def retrieve_operational_top5(
    question: str,
    collection: Any,
    model: Any,
    namespace: dict[str, Any],
    top_k: int = TOP_K,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    if top_k != TOP_K:
        raise RuntimeError(f"평가 Top K는 {TOP_K}로 고정되어 있습니다.")
    intent = namespace["detect_question_intent"](question)
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
        metadata = metadatas[index] if index < len(metadatas) and isinstance(metadatas[index], dict) else {}
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
    selected = namespace["rerank_search_results"](expanded_question, candidates, top_k)
    diagnostics = {
        "question_intent": intent,
        "question_type": question_type,
        "expanded_search_query": expanded_question,
        "internal_count": internal_count,
        "reranking_applied": bool(rerank_config),
    }
    return selected, diagnostics


def safe_metadata(metadata: dict[str, Any]) -> dict[str, Any]:
    allowed = {
        "source",
        "file_name",
        "doc_name",
        "title",
        "source_file",
        "chunk_id",
        "id",
        "doc_id",
        "chunk_index",
        "index",
        "page",
        "page_number",
        "page_start",
        "category",
        "char_count",
        "added_stage",
    }
    result: dict[str, Any] = {}
    for key in sorted(allowed & set(metadata)):
        value = metadata.get(key)
        if isinstance(value, (str, int, float, bool)) or value is None:
            result[key] = str(value)[:500] if isinstance(value, str) else value
    return result


def metadata_page(metadata: dict[str, Any]) -> str:
    for key in ("page", "page_number", "page_start", "source_page", "source_page_number"):
        value = metadata.get(key)
        if value not in (None, ""):
            return str(value)
    return ""


def enrich_retrieval_hit(hit: dict[str, Any], patterns: list[dict[str, Any]], core_terms: list[dict[str, Any]]) -> dict[str, Any]:
    metadata = hit.get("metadata") if isinstance(hit.get("metadata"), dict) else {}
    document_title = str(hit.get("source", "")).strip()
    chunk_id = str(hit.get("chunk_id", "")).strip()
    text = str(hit.get("text", ""))
    document_match = match_expected_document(document_title, patterns)
    matched_terms = []
    term_reasons = []
    for term in core_terms:
        matched, reason = match_core_term(text, term)
        if matched:
            matched_terms.append(term.get("term", ""))
            term_reasons.append(reason)
    return {
        "rank": int(hit.get("rank", 0)),
        "document_title": document_title,
        "chunk_id": chunk_id,
        "source": str(metadata.get("source", "")).strip(),
        "page": metadata_page(metadata),
        "distance": hit.get("distance"),
        "document_text_excerpt": text[:500],
        "document_text": text,
        "metadata_json_safe": json.dumps(safe_metadata(metadata), ensure_ascii=False, separators=(",", ":")),
        "document_match": document_match,
        "matched_core_terms": matched_terms,
        "core_match_reasons": term_reasons,
    }


def compute_document_hit_metrics(hits: list[dict[str, Any]]) -> dict[str, Any]:
    matched_ranks = [hit["rank"] for hit in hits if hit["document_match"]["matched"]]
    first_rank = min(matched_ranks) if matched_ranks else 0
    return {
        "first_expected_document_rank": first_rank,
        "document_hit_at_1": int(bool(first_rank and first_rank <= 1)),
        "document_hit_at_3": int(bool(first_rank and first_rank <= 3)),
        "document_hit_at_5": int(bool(first_rank and first_rank <= 5)),
        "reciprocal_rank": (1.0 / first_rank) if first_rank else 0.0,
    }


def compute_element_coverage(hits: list[dict[str, Any]], core_terms: list[dict[str, Any]], k: int) -> tuple[float, list[str], list[str]]:
    if not core_terms:
        return 0.0, [], []
    merged_text = " ".join(hit.get("document_text", "") for hit in hits[:k])
    matched_terms: list[str] = []
    reasons: list[str] = []
    for term in core_terms:
        matched, reason = match_core_term(merged_text, term)
        if matched:
            matched_terms.append(term.get("term", ""))
            reasons.append(f"{term.get('term', '')}:{reason}")
    return len(matched_terms) / len(core_terms), matched_terms, reasons


def compute_retrieval_diversity(hits: list[dict[str, Any]]) -> dict[str, Any]:
    documents = [hit.get("document_title", "") for hit in hits if hit.get("document_title", "")]
    counts = Counter(documents)
    total = len(documents)
    maximum = max(counts.values(), default=0)
    return {
        "unique_document_count_top5": len(counts),
        "duplicate_document_count_top5": max(0, total - len(counts)),
        "top_document_concentration": maximum,
        "same_document_ratio": (maximum / total) if total else 0.0,
    }


def compute_metadata_completeness(hits: list[dict[str, Any]]) -> float:
    if not hits:
        return 0.0
    total_fields = 0
    present_fields = 0
    for hit in hits:
        fields = (
            hit.get("document_title", ""),
            hit.get("chunk_id", ""),
            hit.get("source", ""),
            hit.get("page", ""),
        )
        total_fields += len(fields)
        present_fields += sum(value not in (None, "", "정보 없음", "출처 정보 없음") for value in fields)
    return present_fields / total_fields if total_fields else 0.0


def decide_auto_review(
    patterns: list[dict[str, Any]],
    metrics: dict[str, Any],
    diversity: dict[str, Any],
    search_error: str = "",
) -> tuple[str, str]:
    reasons = []
    if search_error:
        reasons.append("검색 오류")
    if not patterns:
        reasons.append("기대 문서 패턴 없음")
    if metrics.get("document_hit_at_5", 0) == 0:
        reasons.append("기대 문서 Hit@5=0")
    if metrics.get("element_coverage_at_5", 0.0) < 0.4:
        reasons.append("Element Coverage@5<0.4")
    if (
        diversity.get("unique_document_count_top5", 0) == 1
        and metrics.get("document_hit_at_5", 0) == 0
    ):
        reasons.append("Top5 단일 문서 편중 및 기대 문서 없음")
    if metrics.get("metadata_completeness_top5", 0.0) < 0.75:
        reasons.append("metadata 완전성 낮음")
    if metrics.get("document_hit_at_5", 0) == 0 and metrics.get("element_coverage_at_5", 0.0) >= 0.7:
        reasons.append("문서명 불일치·핵심 요소 고포함")
    if metrics.get("document_hit_at_5", 0) == 1 and metrics.get("element_coverage_at_5", 0.0) < 0.2:
        reasons.append("문서명 일치·핵심 요소 저포함")
    if reasons:
        return "REVIEW", "; ".join(reasons)
    if metrics.get("document_hit_at_3", 0) == 0 or metrics.get("element_coverage_at_3", 0.0) < 0.6:
        return "CHECK", "상위 3개 결과에 대한 추가 확인 권장"
    return "PASS", "자동 규칙 기준 통과(수동 검토 전 1차 평가)"


def mean(values: list[float]) -> float:
    return sum(values) / len(values) if values else 0.0


def aggregate_results(rows: list[dict[str, Any]], group_field: str) -> list[dict[str, Any]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        grouped[str(row[group_field])].append(row)
    summaries = []
    for group_name, group_rows in sorted(grouped.items()):
        summaries.append(
            {
                group_field: group_name,
                "question_count": len(group_rows),
                "document_hit_at_1": mean([float(row["document_hit_at_1"]) for row in group_rows]),
                "document_hit_at_3": mean([float(row["document_hit_at_3"]) for row in group_rows]),
                "document_hit_at_5": mean([float(row["document_hit_at_5"]) for row in group_rows]),
                "mrr": mean([float(row["reciprocal_rank"]) for row in group_rows]),
                "element_coverage_at_1": mean([float(row["element_coverage_at_1"]) for row in group_rows]),
                "element_coverage_at_3": mean([float(row["element_coverage_at_3"]) for row in group_rows]),
                "element_coverage_at_5": mean([float(row["element_coverage_at_5"]) for row in group_rows]),
                "metadata_completeness": mean([float(row["metadata_completeness_top5"]) for row in group_rows]),
                "review_count": sum(row["auto_review_status"] == "REVIEW" for row in group_rows),
            }
        )
    return summaries


def write_tsv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.write_bytes(serialize_tsv(rows, fieldnames))


def write_charts(
    overall: dict[str, float],
    category_rows: list[dict[str, Any]],
    difficulty_rows: list[dict[str, Any]],
) -> None:
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ModuleNotFoundError:
        write_charts_with_existing_runtime(overall, category_rows, difficulty_rows)
        return

    plt.rcParams["font.family"] = "Malgun Gothic"
    plt.rcParams["axes.unicode_minus"] = False

    labels = ["Hit@1", "Hit@3", "Hit@5", "MRR"]
    values = [overall["hit_at_1"], overall["hit_at_3"], overall["hit_at_5"], overall["mrr"]]
    figure, axis = plt.subplots(figsize=(8, 5))
    bars = axis.bar(labels, values)
    axis.set_title("30문항 기준 RAG 검색 적합성 지표")
    axis.set_xlabel("검색 지표")
    axis.set_ylabel("비율")
    axis.set_ylim(0, 1)
    axis.bar_label(bars, labels=[f"{value:.3f}" for value in values], padding=3)
    figure.tight_layout()
    figure.savefig(HIT_CHART_PATH, dpi=180, bbox_inches="tight")
    plt.close(figure)

    category_labels = [row["normalized_category"] for row in category_rows]
    category_values = [row["document_hit_at_3"] for row in category_rows]
    figure, axis = plt.subplots(figsize=(11, 6))
    bars = axis.barh(category_labels, category_values)
    axis.set_title("질문 유형별 기대 문서 Hit@3")
    axis.set_xlabel("Hit@3")
    axis.set_ylabel("질문 유형")
    axis.set_xlim(0, 1)
    axis.bar_label(bars, labels=[f"{value:.2f}" for value in category_values], padding=3)
    figure.tight_layout()
    figure.savefig(CATEGORY_CHART_PATH, dpi=180, bbox_inches="tight")
    plt.close(figure)

    difficulty_labels = [row["difficulty"] for row in difficulty_rows]
    difficulty_values = [row["document_hit_at_3"] for row in difficulty_rows]
    figure, axis = plt.subplots(figsize=(8, 5))
    bars = axis.bar(difficulty_labels, difficulty_values)
    axis.set_title("난이도별 기대 문서 Hit@3")
    axis.set_xlabel("난이도")
    axis.set_ylabel("Hit@3")
    axis.set_ylim(0, 1)
    axis.bar_label(bars, labels=[f"{value:.3f}" for value in difficulty_values], padding=3)
    figure.tight_layout()
    figure.savefig(DIFFICULTY_CHART_PATH, dpi=180, bbox_inches="tight")
    plt.close(figure)


MATPLOTLIB_CHART_SOURCE = r'''
import json
from pathlib import Path
import sys
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

payload = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
output_dir = Path(sys.argv[2])
plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False

labels = ["Hit@1", "Hit@3", "Hit@5", "MRR"]
values = [payload["overall"]["hit_at_1"], payload["overall"]["hit_at_3"], payload["overall"]["hit_at_5"], payload["overall"]["mrr"]]
figure, axis = plt.subplots(figsize=(8, 5))
bars = axis.bar(labels, values)
axis.set_title("30문항 기준 RAG 검색 적합성 지표")
axis.set_xlabel("검색 지표")
axis.set_ylabel("비율")
axis.set_ylim(0, 1)
axis.bar_label(bars, labels=[f"{value:.3f}" for value in values], padding=3)
figure.tight_layout()
figure.savefig(output_dir / "rag_retrieval_hit_at_k.png", dpi=180, bbox_inches="tight")
plt.close(figure)

category_labels = [row["normalized_category"] for row in payload["category_rows"]]
category_values = [row["document_hit_at_3"] for row in payload["category_rows"]]
figure, axis = plt.subplots(figsize=(11, 6))
bars = axis.barh(category_labels, category_values)
axis.set_title("질문 유형별 기대 문서 Hit@3")
axis.set_xlabel("Hit@3")
axis.set_ylabel("질문 유형")
axis.set_xlim(0, 1)
axis.bar_label(bars, labels=[f"{value:.2f}" for value in category_values], padding=3)
figure.tight_layout()
figure.savefig(output_dir / "rag_retrieval_category_hit3.png", dpi=180, bbox_inches="tight")
plt.close(figure)

difficulty_labels = [row["difficulty"] for row in payload["difficulty_rows"]]
difficulty_values = [row["document_hit_at_3"] for row in payload["difficulty_rows"]]
figure, axis = plt.subplots(figsize=(8, 5))
bars = axis.bar(difficulty_labels, difficulty_values)
axis.set_title("난이도별 기대 문서 Hit@3")
axis.set_xlabel("난이도")
axis.set_ylabel("Hit@3")
axis.set_ylim(0, 1)
axis.bar_label(bars, labels=[f"{value:.3f}" for value in difficulty_values], padding=3)
figure.tight_layout()
figure.savefig(output_dir / "rag_retrieval_difficulty_hit3.png", dpi=180, bbox_inches="tight")
plt.close(figure)
'''


def write_charts_with_existing_runtime(
    overall: dict[str, float],
    category_rows: list[dict[str, Any]],
    difficulty_rows: list[dict[str, Any]],
) -> None:
    runtime_root = Path.home() / "AppData" / "Local" / "spyder-6" / "envs" / "spyder-runtime"
    python_executable = runtime_root / "python.exe"
    if not python_executable.is_file():
        raise RuntimeError("matplotlib이 설치된 기존 로컬 Python runtime을 찾을 수 없습니다.")
    with tempfile.TemporaryDirectory(prefix="minesafe_rag_eval_charts_") as temporary:
        temporary_dir = Path(temporary)
        payload_path = temporary_dir / "chart_payload.json"
        script_path = temporary_dir / "render_charts.py"
        payload_path.write_text(
            json.dumps(
                {
                    "overall": overall,
                    "category_rows": category_rows,
                    "difficulty_rows": difficulty_rows,
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        script_path.write_text(textwrap.dedent(MATPLOTLIB_CHART_SOURCE), encoding="utf-8")
        child_path = ";".join(
            [str(runtime_root), str(runtime_root / "Library" / "bin"), str(runtime_root / "Scripts")]
        )
        command = (
            f"$env:PATH='{child_path};' + $env:PATH; "
            f"$env:PYTHONHOME='{runtime_root}'; "
            f"$env:MPLCONFIGDIR='{temporary_dir}'; "
            f"& '{python_executable}' -B '{script_path}' '{payload_path}' '{OUTPUT_DIR}'"
        )
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        if completed.returncode != 0:
            raise RuntimeError(f"기존 matplotlib runtime 그래프 생성 실패: {completed.stderr[-1200:]}")


ARTIFACT_BUILDER_SOURCE = r'''
import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const payload = JSON.parse(await fs.readFile(process.argv[2], "utf8"));
const outputDir = process.argv[3];

function widthForHeader(header) {
  const text = String(header);
  if (/question|expected|reason|excerpt|comment|주의|방법|설명|문구|문서/.test(text)) return 34;
  if (/metadata|patterns|terms|sha256/.test(text)) return 28;
  if (/category|분류|유형/.test(text)) return 23;
  if (/eval_id|rank|count|Hit|MRR|coverage|완전성|비율|난이도/.test(text)) return 14;
  return 18;
}

function addTableSheet(workbook, name, headers, rows, options = {}) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const safeRows = rows.length ? rows : [Object.fromEntries(headers.map((header) => [header, null]))];
  const values = [headers, ...safeRows.map((row) => headers.map((header) => row[header] ?? null))];
  const range = sheet.getRangeByIndexes(0, 0, values.length, headers.length);
  range.values = values;
  range.format = {
    font: { name: "Malgun Gothic", size: 10, color: "#111827" },
    verticalAlignment: "top",
    wrapText: true,
  };
  const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
  headerRange.format = {
    fill: "#1F4E78",
    font: { name: "Malgun Gothic", size: 10, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    rowHeight: 30,
    borders: { preset: "outside", style: "thin", color: "#9CA3AF" },
  };
  sheet.freezePanes.freezeRows(1);
  for (let index = 0; index < headers.length; index += 1) {
    const columnRange = sheet.getRangeByIndexes(0, index, values.length, 1);
    columnRange.format.columnWidth = widthForHeader(headers[index]);
    if (/hit_at|coverage|reciprocal|mrr|completeness|ratio/.test(String(headers[index]).toLowerCase())) {
      if (values.length > 1) sheet.getRangeByIndexes(1, index, values.length - 1, 1).format.numberFormat = "0.000";
    }
  }
  if (values.length > 1) {
    sheet.getRangeByIndexes(1, 0, values.length - 1, headers.length).format.rowHeight = options.rowHeight ?? 38;
  }
  if (options.validation && values.length > 1) {
    for (const validation of options.validation) {
      const index = headers.indexOf(validation.header);
      if (index >= 0) {
        sheet.getRangeByIndexes(1, index, values.length - 1, 1).dataValidation = {
          rule: { type: "list", values: validation.values },
        };
      }
    }
  }
  return sheet;
}

function addKeyValueSheet(workbook, name, rows) {
  return addTableSheet(workbook, name, ["항목", "값"], rows, { rowHeight: 28 });
}

const results = Workbook.create();
addKeyValueSheet(results, "00_요약", payload.summaryRows);
addTableSheet(results, "01_30문항_질문세트", payload.questionHeaders, payload.questionRows);
addTableSheet(results, "02_질문별_검색결과", payload.resultHeaders, payload.resultRows);
addTableSheet(results, "03_Top5_상세", payload.detailHeaders, payload.detailRows);
addTableSheet(results, "04_유형별_결과", payload.categoryHeaders, payload.categoryRows);
addTableSheet(results, "05_난이도별_결과", payload.difficultyHeaders, payload.difficultyRows);
addTableSheet(results, "06_문서별_검색빈도", payload.frequencyHeaders, payload.frequencyRows);
addTableSheet(results, "07_핵심요소_포함률", payload.elementHeaders, payload.elementRows);
addTableSheet(results, "08_수동검토_대상", payload.reviewHeaders, payload.reviewRows);
addTableSheet(results, "09_평가방법_주의사항", ["구분", "내용"], payload.methodRows, { rowHeight: 44 });
const resultsFile = await SpreadsheetFile.exportXlsx(results);
await resultsFile.save(path.join(outputDir, "rag_retrieval_eval_results_30.xlsx"));

const manual = Workbook.create();
addTableSheet(
  manual,
  "수동검토",
  payload.manualHeaders,
  payload.manualRows,
  {
    rowHeight: 54,
    validation: [
      { header: "human_relevance", values: ["2", "1", "0"] },
      { header: "human_document_match", values: ["Y", "N", "REVIEW"] },
    ],
  },
);
const manualFile = await SpreadsheetFile.exportXlsx(manual);
await manualFile.save(path.join(outputDir, "rag_retrieval_manual_review_30.xlsx"));

const summary = Workbook.create();
addKeyValueSheet(summary, "전체요약", payload.summaryRows);
addTableSheet(summary, "유형별", payload.categoryHeaders, payload.categoryRows);
addTableSheet(summary, "난이도별", payload.difficultyHeaders, payload.difficultyRows);
addTableSheet(summary, "문서별검색빈도", payload.frequencyHeaders, payload.frequencyRows);
addTableSheet(summary, "평가해석", ["지표", "해석"], payload.interpretationRows, { rowHeight: 44 });
const summaryFile = await SpreadsheetFile.exportXlsx(summary);
await summaryFile.save(path.join(outputDir, "rag_retrieval_summary.xlsx"));
'''


def artifact_runtime_paths() -> tuple[Path, Path]:
    runtime_root = Path.home() / ".cache" / "codex-runtimes" / "codex-primary-runtime" / "dependencies"
    node_executable = runtime_root / "node" / "bin" / "node.exe"
    node_modules = runtime_root / "node" / "node_modules"
    if not node_executable.is_file() or not node_modules.is_dir():
        raise RuntimeError("Codex bundled artifact-tool runtime을 찾을 수 없습니다.")
    return node_executable, node_modules


def strip_xml_invalid_characters(value: str) -> str:
    """Remove only characters forbidden by XML 1.0 workbook serialization."""
    return "".join(
        character
        for character in value
        if character in "\t\n\r"
        or "\x20" <= character <= "\ud7ff"
        or "\ue000" <= character <= "\ufffd"
        or "\U00010000" <= character <= "\U0010ffff"
    )


def sanitize_workbook_payload(value: Any) -> Any:
    """Recursively sanitize strings at the Excel export boundary only."""
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


def build_workbooks_with_artifact_tool(payload: dict[str, Any], output_dir: Path) -> None:
    node_executable, node_modules = artifact_runtime_paths()
    with tempfile.TemporaryDirectory(prefix="minesafe_rag_eval_xlsx_") as temporary:
        temporary_dir = Path(temporary)
        junction = temporary_dir / "node_modules"
        linked = subprocess.run(
            ["cmd", "/c", "mklink", "/J", str(junction), str(node_modules)],
            check=False,
            capture_output=True,
            text=True,
        )
        if linked.returncode != 0 or not junction.exists():
            raise RuntimeError("artifact-tool node_modules junction 생성 실패")
        payload_path = temporary_dir / "payload.json"
        builder_path = temporary_dir / "build_workbooks.mjs"
        workbook_payload = sanitize_workbook_payload(payload)
        payload_path.write_text(json.dumps(workbook_payload, ensure_ascii=False), encoding="utf-8")
        builder_path.write_text(textwrap.dedent(ARTIFACT_BUILDER_SOURCE), encoding="utf-8")
        completed = subprocess.run(
            [str(node_executable), str(builder_path), str(payload_path), str(output_dir)],
            cwd=temporary_dir,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        if completed.returncode != 0:
            raise RuntimeError(f"artifact-tool workbook 생성 실패: {completed.stderr[-1200:]}")
        # The bundled artifact runtime may emit inspection sidecars next to exports.
        # They are transient diagnostics, not part of the frozen evaluation outputs.
        for workbook_path in (RESULT_XLSX_PATH, MANUAL_XLSX_PATH, SUMMARY_XLSX_PATH):
            sidecar_path = Path(f"{workbook_path}.inspect.ndjson")
            if sidecar_path.is_file():
                sidecar_path.unlink()


def make_workbook_payload(
    gold_rows: list[dict[str, Any]],
    result_rows: list[dict[str, Any]],
    detail_rows: list[dict[str, Any]],
    category_rows: list[dict[str, Any]],
    difficulty_rows: list[dict[str, Any]],
    frequency_rows: list[dict[str, Any]],
    element_rows: list[dict[str, Any]],
    summary_rows: list[dict[str, Any]],
) -> dict[str, Any]:
    manual_headers = [
        "eval_id",
        "source_question_id",
        "category",
        "difficulty",
        "question",
        "expected_document",
        "expected_core_elements",
        "rank",
        "retrieved_document",
        "chunk_id",
        "excerpt",
        "automatic_document_match",
        "automatic_element_match",
        "human_relevance",
        "human_document_match",
        "human_comment",
    ]
    manual_rows = []
    for row in detail_rows:
        manual_rows.append(
            {
                "eval_id": row["eval_id"],
                "source_question_id": row["source_question_id"],
                "category": row["normalized_category"],
                "difficulty": row["difficulty"],
                "question": row["question"],
                "expected_document": row["expected_document_raw"],
                "expected_core_elements": row["expected_core_elements_raw"],
                "rank": row["rank"],
                "retrieved_document": row["document_title"],
                "chunk_id": row["chunk_id"],
                "excerpt": row["document_text_excerpt"],
                "automatic_document_match": row["automatic_document_match"],
                "automatic_element_match": row["automatic_element_match"],
                "human_relevance": None,
                "human_document_match": None,
                "human_comment": None,
            }
        )
    method_rows = [
        {"구분": "평가 종류", "내용": "30문항 기준 RAG 검색 적합성 평가이며 기능 회귀 테스트·최종 답변 정확성 평가와 구분한다."},
        {"구분": "문서 지표", "내용": "기대 문서 기준 Hit@1/3/5와 MRR을 계산한다."},
        {"구분": "핵심 요소", "내용": "검색 chunk에서 기대 안전조치 용어가 나타나는지를 규칙 기반으로 계산한다."},
        {"구분": "distance", "내용": "ChromaDB distance는 진단값으로만 기록하며 정확도 확률로 변환하지 않는다."},
        {"구분": "해석 제한", "내용": "수동 검토 전 1차 평가이며 현재 평가 세트와 문서 구성에 따른 결과다."},
        {"구분": "법적 판단", "내용": "법적 정확도나 전체 답변 정확도를 의미하지 않는다."},
    ]
    interpretation_rows = [
        {"지표": "Hit@3", "해석": "30문항 중 기대 문서가 상위 3개 검색 결과 안에 포함된 비율"},
        {"지표": "MRR", "해석": "기대 문서가 검색 결과의 앞 순위에 배치되는 정도"},
        {"지표": "Element Coverage@K", "해석": "기대 안전조치 용어가 상위 K개 chunk에 포함된 규칙 기반 보조 지표"},
        {"지표": "metadata 완전성", "해석": "문서명·chunk_id·source·page 정보가 존재하는 비율"},
        {"지표": "주의", "해석": "어떤 지표도 법적 정확도 또는 최종 답변 정확도가 아님"},
    ]
    detail_headers = [
        "eval_id", "source_question_id", "normalized_category", "difficulty", "question",
        "expected_document_raw", "expected_core_elements_raw", "rank", "document_title",
        "chunk_id", "source", "page", "distance", "document_text_excerpt", "metadata_json_safe",
        "automatic_document_match", "exact_normalized_match", "title_token_match",
        "matched_expected_pattern", "match_reason", "automatic_element_match", "core_match_reasons",
    ]
    category_headers = list(category_rows[0].keys()) if category_rows else ["normalized_category"]
    difficulty_headers = list(difficulty_rows[0].keys()) if difficulty_rows else ["difficulty"]
    frequency_headers = list(frequency_rows[0].keys()) if frequency_rows else ["document_title", "retrieval_count"]
    element_headers = list(element_rows[0].keys()) if element_rows else ["eval_id"]
    review_rows = [row for row in result_rows if row["auto_review_status"] == "REVIEW"]
    return {
        "summaryRows": summary_rows,
        "questionHeaders": GOLD_COLUMNS,
        "questionRows": gold_rows,
        "resultHeaders": RESULT_COLUMNS,
        "resultRows": result_rows,
        "detailHeaders": detail_headers,
        "detailRows": detail_rows,
        "categoryHeaders": category_headers,
        "categoryRows": category_rows,
        "difficultyHeaders": difficulty_headers,
        "difficultyRows": difficulty_rows,
        "frequencyHeaders": frequency_headers,
        "frequencyRows": frequency_rows,
        "elementHeaders": element_headers,
        "elementRows": element_rows,
        "reviewHeaders": RESULT_COLUMNS,
        "reviewRows": review_rows,
        "methodRows": method_rows,
        "manualHeaders": manual_headers,
        "manualRows": manual_rows,
        "interpretationRows": interpretation_rows,
    }


def report_group_lines(rows: list[dict[str, Any]], label_key: str) -> str:
    return "\n".join(
        f"- {row[label_key]}: n={row['question_count']}, Hit@3={row['document_hit_at_3']:.3f}, "
        f"MRR={row['mrr']:.3f}, Coverage@3={row['element_coverage_at_3']:.3f}"
        for row in rows
    )


def write_quality_report(
    *,
    source_before_sha: str,
    source_after_sha: str,
    app_before_sha: str,
    app_after_sha: str,
    gold_sha: str,
    script_sha: str,
    logical_before: dict[str, Any],
    logical_after: dict[str, Any],
    physical_before: dict[str, Any],
    physical_after: dict[str, Any],
    overall: dict[str, float],
    category_rows: list[dict[str, Any]],
    difficulty_rows: list[dict[str, Any]],
    frequency_rows: list[dict[str, Any]],
    review_ids: list[str],
    failure_ids: list[str],
    category_counts: Counter[str],
    difficulty_counts: Counter[str],
    run_time: str,
) -> None:
    strongest = max(category_rows, key=lambda row: (row["document_hit_at_3"], row["mrr"])) if category_rows else {}
    weakest = min(category_rows, key=lambda row: (row["document_hit_at_3"], row["mrr"])) if category_rows else {}
    top_documents = "\n".join(
        f"- {row['document_title']}: {row['retrieval_count']}회" for row in frequency_rows[:10]
    ) or "- 없음"
    generated_files = "\n".join(
        f"- {path.name}"
        for path in (
            Path(__file__), GOLD_PATH, RESULT_TSV_PATH, RESULT_XLSX_PATH, MANUAL_XLSX_PATH,
            SUMMARY_XLSX_PATH, HIT_CHART_PATH, CATEGORY_CHART_PATH, DIFFICULTY_CHART_PATH, REPORT_PATH,
            TEST_PATH,
        )
    )
    report = f"""MineSafe AI 30문항 RAG 검색 적합성 평가 보고서

1. 연구 목적
기존 110문항 질문 시나리오에서 균형 있게 고정한 30문항을 이용해, 사용자 질문에 적합한 공식 문서가 현재 운영 RAG 검색 결과의 상위 순위에 포함되는지 평가한다.

2. 기존 평가와 이번 평가의 차이
기능 회귀 테스트는 코드 구조와 기존 기능 유지를 검사한다. RAG 검색 적합성 평가는 기대 공식 문서가 Top K에 포함되는지를 검사한다. 답변 정확성 평가는 최종 생성 답변의 법적·현장 타당성을 확인하는 별도 연구 대상이다.

3. 110문항 원본 정보
- 파일: 02_질문시나리오/question_scenarios_110.tsv
- 행 수: 110
- 작업 전 SHA-256: {source_before_sha}
- 작업 후 SHA-256: {source_after_sha}

4. 30문항 선정 방식
10개 목표 유형별 3문항을 고정 seed {RANDOM_SEED}와 명시적 고정 선정 계획으로 선택했다. 원본 난이도는 하 1·중 92·상 17이므로 유형별 하·중·상 1개씩 구성할 수 없었다. 컨베이어 직접 문항이 1건뿐이어서 정비 전 에너지 차단·잠금표지 인접 문항 2건을 보충했고 사유를 gold set에 기록했다. 검색 결과를 보기 전에 gold set을 저장·동결했다.

5. 유형별 문항 수
{chr(10).join(f'- {key}: {value}문항' for key, value in sorted(category_counts.items()))}

6. 난이도별 문항 수
{chr(10).join(f'- {key}: {value}문항' for key, value in sorted(difficulty_counts.items()))}

7. gold set SHA-256
{gold_sha}

8. 기존 RAG 설정
- app.py를 import하지 않고 AST로 운영 검색 순수 함수·상수를 추출했다.
- 질문 의도 판정 → 검색어 확장 → normalized embedding → 내부 후보 검색 → 일부 유형의 선호 문서 후보 추가 → 운영 휴리스틱 재정렬·문서 다양화 → Top 5 순서를 그대로 사용했다.
- normalize_embeddings={NORMALIZE_EMBEDDINGS}

9. DB 경로와 collection
- 경로: 10_vector_db_with_major_accident_docs
- collection: {COLLECTION_NAME}

10. collection count
작업 전 {logical_before['count']}개, 작업 후 {logical_after['count']}개다.

11. embedding 모델과 차원
- 모델: {EMBEDDING_MODEL_NAME}
- 차원: {EMBEDDING_DIMENSION}
- sentence-transformers: {importlib.metadata.version('sentence-transformers')}
- ChromaDB: {importlib.metadata.version('chromadb')}

12. Top K
질문별 최종 Top {TOP_K}를 평가했다. 유형별 운영 rerank가 있는 경우 내부 후보 수는 app.py 운영 설정을 따랐다.

13. 기대 문서 매칭 규칙
확장자·선행 번호·공백·밑줄·하이픈을 정규화한 제목 포함 일치와 75% 이상 핵심 제목 토큰 일치를 사용했다. 광산 안전 지침, 위험성평가 안내서, 중대재해처벌법의 공식 제목 alias만 명시적으로 허용했다. 느슨한 fuzzy match는 사용하지 않았다.

14. 핵심 요소 포함률 규칙
원본 핵심 요소를 구분자로 분리하고, 원본에 존재하는 요소에 대해서만 작업중지·전원 차단·잠금표지 등 제한된 표현 alias를 허용했다. 의미 이해가 아닌 규칙 기반 핵심 용어 포함률이다.

15. Document Hit@1
{overall['hit_at_1']:.4f}

16. Document Hit@3
{overall['hit_at_3']:.4f}. 이는 30문항 중 이 비율에서 기대 문서가 상위 3개 검색 결과 안에 포함됐다는 의미이며 법적 정확도를 뜻하지 않는다.

17. Document Hit@5
{overall['hit_at_5']:.4f}

18. MRR
{overall['mrr']:.4f}. 기대 문서가 검색 결과의 앞 순위에 배치되는 정도를 나타낸다.

19. Element Coverage@1
{overall['element_coverage_at_1']:.4f}

20. Element Coverage@3
{overall['element_coverage_at_3']:.4f}

21. Element Coverage@5
{overall['element_coverage_at_5']:.4f}. 기대 안전조치 용어가 검색된 chunk 안에 포함된 비율을 나타내는 규칙 기반 보조 지표다.

22. 유형별 결과
{report_group_lines(category_rows, 'normalized_category')}

23. 난이도별 결과
{report_group_lines(difficulty_rows, 'difficulty')}

24. 문서별 검색 빈도
{top_documents}

25. metadata 완전성
Top 5의 문서명·chunk_id·source·page 필드 기준 평균 {overall['metadata_completeness']:.4f}다. 없는 page 정보는 임의 보완하지 않았다.

26. 수동 검토 필요 문항
{len(review_ids)}문항: {', '.join(review_ids) if review_ids else '없음'}

27. 검색 실패 문항 목록
{len(failure_ids)}문항: {', '.join(failure_ids) if failure_ids else '없음'}

28. 검색이 강한 질문 유형
{strongest.get('normalized_category', '없음')} (Hit@3={strongest.get('document_hit_at_3', 0):.3f}, MRR={strongest.get('mrr', 0):.3f})

29. 검색이 약한 질문 유형
{weakest.get('normalized_category', '없음')} (Hit@3={weakest.get('document_hit_at_3', 0):.3f}, MRR={weakest.get('mrr', 0):.3f})

30. 자동평가의 한계
기대 문서명과 DB 제목 차이, 다른 공식 문서의 직접 근거 가능성, 제목은 맞지만 chunk 내용이 다른 경우, 동의어 표현을 자동 규칙만으로 완전히 판정할 수 없다. V1 핵심 요소 파서에서 가운데점은 원본 요구에 따라 구분자로 동결되어 '잠금·표지' 같은 복합 표현이 분리될 수 있으며, 결과를 본 뒤 gold set을 바꾸지 않기 위해 이 한계를 그대로 기록한다. 결과는 수동 검토 전 1차 평가이며 30문항을 전체 광산 안전 질문으로 일반화하지 않는다.

31. 법적 정확도 지표가 아닌 이유
Hit@K·MRR은 문서 검색 순위를, Element Coverage는 용어 출현을 측정할 뿐 법령 해석·현장 사실관계·최종 답변의 타당성을 평가하지 않는다. ChromaDB distance도 정확도 확률로 변환하지 않았다.

32. 기존 200개 회귀 테스트와의 차이
기존 200개 테스트는 앱 기능·코드 계약의 회귀 여부를 검사했다. 이번 평가는 별도 동결 gold set을 사용해 기대 문서의 실제 검색 순위를 측정한다.

33. 개선이 필요한 검색 영역
유형별 Hit@3와 수동 검토 대상 문항을 기준으로 후속 실험 후보를 정하되, 이번 baseline 실행 중에는 query expansion·Top K·distance·embedding·DB·gold set을 튜닝하지 않았다.

34. 향후 수동 검토 방법
rag_retrieval_manual_review_30.xlsx에서 각 질문의 Top 5를 블라인드로 읽고 human_relevance(2/1/0), human_document_match(Y/N/REVIEW), human_comment를 사람이 직접 입력한다. 자동 칸은 최종 전문가 판정이 아니다.

35. 생성 파일
{generated_files}

36. 기존 파일 무결성 결과
- app.py 작업 전/후 SHA-256: {app_before_sha} / {app_after_sha}
- 원본 질문 작업 전/후 SHA-256: {source_before_sha} / {source_after_sha}
- DB 논리 count: {logical_before['count']} / {logical_after['count']}
- DB ID SHA-256: {logical_before['ids_sha256']} / {logical_after['ids_sha256']}
- DB 물리 metadata 지문: {physical_before['metadata_sha256']} / {physical_after['metadata_sha256']}
Chroma 조회가 SQLite·인덱스 파일의 수정시각을 갱신할 수 있어 물리 metadata 지문은 진단값으로만 기록하고, 논리 count와 ID 목록 동일 여부를 데이터 무결성 기준으로 사용했다.

37. 신규 테스트 결과
평가 산출물 생성 후 별도 명령으로 실행하며 최종 검증 시 이 항목을 갱신한다.

38. 전체 테스트 결과
작업 전 200/200 통과. 신규 테스트 포함 최종 결과는 별도 명령으로 실행한 뒤 이 항목을 갱신한다.

39. Git whitelist
사용자가 지정한 평가 폴더 파일 10개와 18_legal_evidence_features/test_rag_retrieval_quality_evaluation.py만 허용한다.

40. Git 제외 파일
app.py, 질문 원본, 기존 평가 파일, 기존 DB, chunks, 인증정보, 백업 및 관련 없는 변경 파일은 stage하지 않는다.

41. commit SHA
보고서 자체를 포함하는 최종 커밋 SHA는 자기참조 때문에 파일 안에 확정할 수 없으며 최종 완료 보고에 기록한다.

42. push 결과
전체 테스트·무결성·whitelist 검증 후 실행하며 최종 완료 보고에 기록한다.

43. 교수님 설명용 문구
기존 자동 테스트는 앱의 기능과 코드 구조가 유지되는지를 검증하는 회귀 테스트였습니다. 이번에는 별도로 30문항의 검색 적합성 평가 세트를 구축하고, 기존 질문 시나리오에 기록된 기대 검색 문서를 기준으로 관련 공식 문서가 검색 결과 상위 1개, 3개, 5개 안에 포함되는지를 Hit@K와 MRR로 평가했습니다. 또한 기대 안전조치 용어가 검색된 문서에 포함되는지를 규칙 기반 핵심 요소 포함률로 분석했습니다. 이 결과는 법적 정확도나 최종 답변 정확도가 아니라 현재 공식 문서 RAG의 검색 성능을 나타내는 평가입니다.

재현성 정보
- 실행 시각: {run_time}
- Python: {platform.python_version()}
- ChromaDB: {importlib.metadata.version('chromadb')}
- sentence-transformers: {importlib.metadata.version('sentence-transformers')}
- embedding 모델: {EMBEDDING_MODEL_NAME}
- embedding 차원: {EMBEDDING_DIMENSION}
- normalize_embeddings: {NORMALIZE_EMBEDDINGS}
- DB 경로: {DB_PATH}
- collection: {COLLECTION_NAME}
- collection count: {logical_after['count']}
- Top K: {TOP_K}
- 원본 SHA-256: {source_after_sha}
- gold set SHA-256: {gold_sha}
- 평가 스크립트 SHA-256: {script_sha}
- random seed: {RANDOM_SEED}

[INTEGRITY]
source_questions_before_sha256={source_before_sha}
source_questions_after_sha256={source_after_sha}
app_before_sha256={app_before_sha}
app_after_sha256={app_after_sha}
db_count_before={logical_before['count']}
db_count_after={logical_after['count']}
db_ids_sha256_before={logical_before['ids_sha256']}
db_ids_sha256_after={logical_after['ids_sha256']}
db_tree_sha256_before={physical_before['metadata_sha256']}
db_tree_sha256_after={physical_after['metadata_sha256']}
gold_set_sha256={gold_sha}
evaluation_script_sha256={script_sha}
"""
    REPORT_PATH.write_text(report, encoding="utf-8")


def update_report_validation_status(new_test_result: str, full_test_result: str) -> None:
    report = REPORT_PATH.read_text(encoding="utf-8")
    report = re.sub(
        r"(37\. 신규 테스트 결과\n).*?(?=\n38\. 전체 테스트 결과)",
        lambda match: f"{match.group(1)}{new_test_result}\n",
        report,
        flags=re.DOTALL,
    )
    report = re.sub(
        r"(38\. 전체 테스트 결과\n).*?(?=\n39\. Git whitelist)",
        lambda match: f"{match.group(1)}{full_test_result}\n",
        report,
        flags=re.DOTALL,
    )
    REPORT_PATH.write_text(report, encoding="utf-8")


def run_evaluation() -> dict[str, Any]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    source_before_sha = sha256_file(SOURCE_QUESTIONS_PATH)
    app_before_sha = sha256_file(APP_PATH)
    if source_before_sha != EXPECTED_SOURCE_SHA256:
        raise RuntimeError("원본 110문항 SHA-256이 기준과 다릅니다.")
    if app_before_sha != EXPECTED_APP_SHA256:
        raise RuntimeError("app.py SHA-256이 작업 시작 기준과 다릅니다.")

    source_rows = read_scenario_rows(SOURCE_QUESTIONS_PATH)
    gold_rows = select_balanced_questions(source_rows, RANDOM_SEED)
    gold_bytes = serialize_tsv(gold_rows, GOLD_COLUMNS)
    gold_sha = freeze_or_validate_gold_set(GOLD_PATH, gold_bytes)

    namespace = load_operational_search_namespace(APP_PATH)
    validate_operational_contract(namespace)
    physical_before = directory_metadata_snapshot(DB_PATH)

    import chromadb

    client = chromadb.PersistentClient(path=str(DB_PATH))
    collection = client.get_collection(name=COLLECTION_NAME)
    logical_before = snapshot_collection_logical(collection)
    if logical_before["count"] != 820 or logical_before["id_count"] != 820:
        raise RuntimeError(f"collection count가 820이 아닙니다: {logical_before['count']}")
    peek = collection.peek(limit=1)
    embeddings = peek.get("embeddings")
    if embeddings is None or len(embeddings) == 0 or len(embeddings[0]) != EMBEDDING_DIMENSION:
        raise RuntimeError("collection embedding 차원이 384가 아닙니다.")
    model = load_local_embedding_model(EMBEDDING_MODEL_NAME)

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
            raw_hits, diagnostics = retrieve_operational_top5(
                gold["question"], collection, model, namespace, TOP_K
            )
            hits = [enrich_retrieval_hit(hit, patterns, core_terms) for hit in raw_hits]
        except Exception as error:
            search_error = f"{type(error).__name__}: {str(error)[:200]}"
            failure_ids.append(gold["eval_id"])
            hits = []

        document_metrics = compute_document_hit_metrics(hits)
        coverage_1, matched_1, reasons_1 = compute_element_coverage(hits, core_terms, 1)
        coverage_3, matched_3, reasons_3 = compute_element_coverage(hits, core_terms, 3)
        coverage_5, matched_5, reasons_5 = compute_element_coverage(hits, core_terms, 5)
        diversity = compute_retrieval_diversity(hits)
        metadata_completeness = compute_metadata_completeness(hits)
        metrics = {
            **document_metrics,
            "element_coverage_at_1": coverage_1,
            "element_coverage_at_3": coverage_3,
            "element_coverage_at_5": coverage_5,
            "metadata_completeness_top5": metadata_completeness,
        }
        auto_status, auto_reason = decide_auto_review(patterns, metrics, diversity, search_error)
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
            "gold_set_sha256": gold_sha,
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

    write_tsv(RESULT_TSV_PATH, result_rows, RESULT_COLUMNS)
    category_rows = [
        {
            "normalized_category": row["category"],
            **{key: value for key, value in row.items() if key != "category"},
        }
        for row in aggregate_results(result_rows, "category")
    ]
    difficulty_rows = aggregate_results(result_rows, "difficulty")
    frequency_counter = Counter(row["document_title"] for row in detail_rows if row["document_title"])
    frequency_rows = [
        {"document_title": title, "retrieval_count": count}
        for title, count in frequency_counter.most_common()
    ]
    overall = {
        "hit_at_1": mean([float(row["document_hit_at_1"]) for row in result_rows]),
        "hit_at_3": mean([float(row["document_hit_at_3"]) for row in result_rows]),
        "hit_at_5": mean([float(row["document_hit_at_5"]) for row in result_rows]),
        "mrr": mean([float(row["reciprocal_rank"]) for row in result_rows]),
        "element_coverage_at_1": mean([float(row["element_coverage_at_1"]) for row in result_rows]),
        "element_coverage_at_3": mean([float(row["element_coverage_at_3"]) for row in result_rows]),
        "element_coverage_at_5": mean([float(row["element_coverage_at_5"]) for row in result_rows]),
        "metadata_completeness": mean([float(row["metadata_completeness_top5"]) for row in result_rows]),
    }
    run_time = datetime.now().astimezone().isoformat(timespec="seconds")
    review_ids = [row["eval_id"] for row in result_rows if row["auto_review_status"] == "REVIEW"]
    summary_rows = [
        {"항목": "평가 문항 수", "값": 30},
        {"항목": "DB collection count", "값": logical_before["count"]},
        {"항목": "embedding 모델", "값": EMBEDDING_MODEL_NAME},
        {"항목": "embedding 차원", "값": EMBEDDING_DIMENSION},
        {"항목": "normalize_embeddings", "값": str(NORMALIZE_EMBEDDINGS)},
        {"항목": "Hit@1", "값": overall["hit_at_1"]},
        {"항목": "Hit@3", "값": overall["hit_at_3"]},
        {"항목": "Hit@5", "값": overall["hit_at_5"]},
        {"항목": "MRR", "값": overall["mrr"]},
        {"항목": "평균 Element Coverage@1", "값": overall["element_coverage_at_1"]},
        {"항목": "평균 Element Coverage@3", "값": overall["element_coverage_at_3"]},
        {"항목": "평균 Element Coverage@5", "값": overall["element_coverage_at_5"]},
        {"항목": "metadata 완전성", "값": overall["metadata_completeness"]},
        {"항목": "수동 검토 필요 문항 수", "값": len(review_ids)},
        {"항목": "검색 실패 문항 수", "값": len(failure_ids)},
        {"항목": "gold set SHA-256", "값": gold_sha},
        {"항목": "평가 실행 시각", "값": run_time},
    ]

    write_charts(overall, category_rows, difficulty_rows)
    payload = make_workbook_payload(
        gold_rows,
        result_rows,
        detail_rows,
        category_rows,
        difficulty_rows,
        frequency_rows,
        element_rows,
        summary_rows,
    )
    build_workbooks_with_artifact_tool(payload, OUTPUT_DIR)

    logical_after = snapshot_collection_logical(collection)
    del model, collection, client
    gc.collect()
    physical_after = directory_metadata_snapshot(DB_PATH)
    source_after_sha = sha256_file(SOURCE_QUESTIONS_PATH)
    app_after_sha = sha256_file(APP_PATH)
    if source_after_sha != source_before_sha or app_after_sha != app_before_sha:
        raise RuntimeError("원본 질문 또는 app.py 무결성 오류")
    if (
        logical_after["count"] != logical_before["count"]
        or logical_after["ids_sha256"] != logical_before["ids_sha256"]
    ):
        raise RuntimeError("Vector DB 논리 데이터가 평가 전후 달라졌습니다.")
    script_sha = sha256_file(Path(__file__))
    category_counts = Counter(row["normalized_category"] for row in gold_rows)
    difficulty_counts = Counter(row["difficulty"] for row in gold_rows)
    write_quality_report(
        source_before_sha=source_before_sha,
        source_after_sha=source_after_sha,
        app_before_sha=app_before_sha,
        app_after_sha=app_after_sha,
        gold_sha=gold_sha,
        script_sha=script_sha,
        logical_before=logical_before,
        logical_after=logical_after,
        physical_before=physical_before,
        physical_after=physical_after,
        overall=overall,
        category_rows=category_rows,
        difficulty_rows=difficulty_rows,
        frequency_rows=frequency_rows,
        review_ids=review_ids,
        failure_ids=failure_ids,
        category_counts=category_counts,
        difficulty_counts=difficulty_counts,
        run_time=run_time,
    )
    return {
        "gold_sha256": gold_sha,
        "overall": overall,
        "review_count": len(review_ids),
        "failure_count": len(failure_ids),
        "db_count": logical_after["count"],
        "db_ids_sha256": logical_after["ids_sha256"],
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="MineSafe AI 30문항 RAG 검색 적합성 평가")
    parser.add_argument("--update-validation-only", action="store_true")
    parser.add_argument("--new-test-result", default="")
    parser.add_argument("--full-test-result", default="")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.update_validation_only:
        if not args.new_test_result or not args.full_test_result:
            raise RuntimeError("검증 결과 문자열이 필요합니다.")
        update_report_validation_status(args.new_test_result, args.full_test_result)
        print("report validation status updated")
        return 0
    summary = run_evaluation()
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
