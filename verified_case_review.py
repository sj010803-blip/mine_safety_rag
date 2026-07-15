"""Manual verification workflow for official accident cases.

This module never rewrites the source OCR JSONL and never marks OCR output as
verified automatically.  A verified state can only be created by an explicit
review action that confirms the source image and required fields.
"""

from __future__ import annotations

import csv
import gc
import io
import json
import re
import shutil
import subprocess
import tempfile
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any


ROOT_DIR = Path(__file__).resolve().parent
SOURCE_CASE_DB_DIR = ROOT_DIR / "23_official_accident_case_vector_db"
VERIFIED_CASE_DB_DIR = ROOT_DIR / "23_verified_official_accident_case_vector_db"
VERIFIED_CASE_DB_CANDIDATE_DIR = ROOT_DIR / "23_verified_official_accident_case_vector_db_candidate"
AUTO_SCREENED_CASE_DB_DIR = ROOT_DIR / "23_auto_screened_official_accident_case_vector_db"
AUTO_SCREENED_CASE_DB_CANDIDATE_DIR = (
    ROOT_DIR / "23_auto_screened_official_accident_case_vector_db_candidate"
)
REVIEW_DIR = ROOT_DIR / "25_verified_case_review"
REVIEW_STATUS_PATH = REVIEW_DIR / "official_case_review_status.jsonl"
AUTO_SCREENING_RESULT_PATH = REVIEW_DIR / "auto_screened_quality_results.jsonl"
CARD_IMAGE_DIR = REVIEW_DIR / "card_images"
SOURCE_PDF_DIR = ROOT_DIR / "21_official_accident_case_docs"
ORIGINAL_OCR_JSONL_PATH = (
    ROOT_DIR
    / "21_official_accident_case_ocr"
    / "page_text"
    / "official_siren_ocr_pages.jsonl"
)

LAW_COLLECTION_NAME = "mine_safety_docs"
SOURCE_CASE_COLLECTION_NAME = "mine_official_accident_cases"
VERIFIED_CASE_COLLECTION_NAME = "mine_verified_official_accident_cases"
AUTO_SCREENED_CASE_COLLECTION_NAME = "mine_auto_screened_official_accident_cases"
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"

VERIFICATION_STATUSES = ("unverified", "verified", "rejected", "manual_review")
AUTO_SCREENED_STATUS = "auto_screened"
CASE_STATUSES = (*VERIFICATION_STATUSES, AUTO_SCREENED_STATUS)
PUBLIC_CASE_STATUSES = ("verified", AUTO_SCREENED_STATUS)
DEFAULT_VERIFICATION_STATUS = "unverified"
REQUIRED_VERIFICATION_CHECKS = (
    "summary_matches_source",
    "date_matches_source",
    "industry_matches_source",
    "accident_type_matches_source",
    "single_accident_only",
    "no_ocr_noise",
)
EDITABLE_FIELDS = (
    "accident_date",
    "industry",
    "accident_type",
    "accident_summary",
    "cause_summary",
    "prevention_summary",
)
PRIORITY_GROUP_LIMIT = 3
AUTO_SCREENED_MIN_TEXT_QUALITY = 80
AUTO_SCREENED_MIN_READING_ORDER = 80
AUTO_SCREENED_MIN_METADATA_QUALITY = 70
AUTO_SCREENED_MIN_DISPLAY_CHARS = 40
AUTO_SCREENED_MAX_DISPLAY_CHARS = 450
SAFE_ENGLISH_TOKENS = {
    "TBM", "PPE", "CCTV", "LOTO", "LED", "SIF", "CO", "O2", "CH4", "LPG",
}
KNOWN_OCR_NOISE_TOKENS = {"ZOOL", "XSF", "SCHRHON", "PUWBCT"}
PUBLIC_RANK_ORDER = {
    ("direct", "verified"): 0,
    ("direct", AUTO_SCREENED_STATUS): 1,
    ("analogous", "verified"): 2,
    ("analogous", AUTO_SCREENED_STATUS): 3,
}


class ReviewWorkflowBlocked(RuntimeError):
    """Raised when a verification safety condition is not met."""


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _safe_metadata_value(value: Any) -> str | int | float | bool:
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def default_review_record(case: dict[str, Any]) -> dict[str, Any]:
    """Add manual-review fields without changing the source case_id."""
    record = dict(case)
    record.update(
        {
            "verification_status": DEFAULT_VERIFICATION_STATUS,
            "verified_at": "",
            "verification_note": "",
            "verified_fields": [],
            "original_page_image": "",
            "original_page_number": case.get("page_start", ""),
            "source_document": case.get("source_document", ""),
            "extraction_engine": "tesseract_legacy_unverified",
            "extraction_quality": "unverified",
            "rejection_reason": "",
            "layout_ocr_text": "",
            "review_updated_at": "",
        }
    )
    return record


def load_source_case_records() -> list[dict[str, Any]]:
    """Read the preserved source case DB; never mutate it."""
    if not SOURCE_CASE_DB_DIR.is_dir():
        return []
    try:
        import chromadb

        client = chromadb.PersistentClient(path=str(SOURCE_CASE_DB_DIR))
        collection = client.get_collection(name=SOURCE_CASE_COLLECTION_NAME)
        payload = collection.get(include=["documents", "metadatas"])
    except Exception as error:
        raise ReviewWorkflowBlocked("기존 미검증 사례 DB를 읽을 수 없습니다.") from error
    records: list[dict[str, Any]] = []
    documents = payload.get("documents", []) or []
    metadatas = payload.get("metadatas", []) or []
    for index, case_id in enumerate(payload.get("ids", []) or []):
        metadata = metadatas[index] if index < len(metadatas) else {}
        document = documents[index] if index < len(documents) else ""
        if not isinstance(metadata, dict):
            continue
        record = default_review_record({**metadata, "case_id": str(case_id), "text": str(document or "")})
        records.append(record)
    return sorted(records, key=lambda item: str(item.get("case_id", "")))


def _append_review_event(record: dict[str, Any], event_type: str) -> None:
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    event = dict(record)
    event["event_type"] = event_type
    event["event_at"] = _now_iso()
    with REVIEW_STATUS_PATH.open("a", encoding="utf-8", newline="\n") as output:
        output.write(json.dumps(event, ensure_ascii=False, sort_keys=True) + "\n")


def initialize_review_store() -> list[dict[str, Any]]:
    """Initialize all preserved cases as unverified exactly once."""
    if REVIEW_STATUS_PATH.is_file():
        return load_review_records()
    records = load_source_case_records()
    if not records:
        return []
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    temporary_path = REVIEW_STATUS_PATH.with_suffix(".jsonl.tmp")
    initialized_at = _now_iso()
    with temporary_path.open("w", encoding="utf-8", newline="\n") as output:
        for record in records:
            event = dict(record)
            event.update(
                {
                    "event_type": "initialized_unverified",
                    "event_at": initialized_at,
                    "review_updated_at": initialized_at,
                }
            )
            output.write(json.dumps(event, ensure_ascii=False, sort_keys=True) + "\n")
    temporary_path.replace(REVIEW_STATUS_PATH)
    return records


def load_review_records() -> list[dict[str, Any]]:
    """Return only the latest event for each case_id while preserving history on disk."""
    if not REVIEW_STATUS_PATH.is_file():
        return []
    latest: dict[str, dict[str, Any]] = {}
    with REVIEW_STATUS_PATH.open("r", encoding="utf-8") as input_file:
        for line in input_file:
            if not line.strip():
                continue
            event = json.loads(line)
            case_id = str(event.get("case_id", "")).strip()
            if case_id:
                latest[case_id] = event
    return sorted(latest.values(), key=lambda item: str(item.get("case_id", "")))


def review_status_counts(records: list[dict[str, Any]]) -> dict[str, int]:
    counts = Counter(str(record.get("verification_status", DEFAULT_VERIFICATION_STATUS)) for record in records)
    return {
        "total": len(records),
        "unverified": counts.get("unverified", 0),
        "verified": counts.get("verified", 0),
        "auto_screened": counts.get(AUTO_SCREENED_STATUS, 0),
        "rejected": counts.get("rejected", 0),
        "manual_review": counts.get("manual_review", 0),
    }


def priority_review_candidates(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Select at most three cases for each presentation topic."""
    groups = (
        ("컨베이어 끼임", ("컨베이어", "끼임", "말림", "회전체"), ("끼임", "말림")),
        ("차량 후진·충돌", ("후진", "차량", "지게차", "덤프트럭", "충돌", "깔림"), ("충돌", "깔림")),
        ("낙반·붕괴", ("낙반", "붕괴", "매몰", "무너짐", "암석"), ("붕괴", "매몰", "맞음")),
        ("전기·감전", ("전기", "전원", "감전", "누전"), ("감전",)),
    )
    selected: list[dict[str, Any]] = []
    used: set[str] = set()
    for group_name, keywords, accident_types in groups:
        candidates: list[tuple[int, dict[str, Any]]] = []
        for record in records:
            case_id = str(record.get("case_id", ""))
            if not case_id or case_id in used or record.get("verification_status") == "rejected":
                continue
            accident_type = str(record.get("accident_type", ""))
            text = " ".join(
                str(record.get(field, "") or "")
                for field in ("accident_type", "accident_summary", "cause_summary", "prevention_summary", "equipment")
            )
            score = (3 if accident_type in accident_types else 0) + sum(keyword in text for keyword in keywords)
            if score > 0:
                candidates.append((score, record))
        candidates.sort(key=lambda item: (-item[0], str(item[1].get("case_id", ""))))
        for _score, record in candidates[:PRIORITY_GROUP_LIMIT]:
            item = dict(record)
            item["priority_review_group"] = group_name
            selected.append(item)
            used.add(str(record.get("case_id", "")))
    return selected


def sanitize_display_text(text: Any) -> str:
    """Normalize whitespace only; never infer or rewrite damaged OCR wording."""
    value = str(text or "").replace("\x00", " ").replace("\r", "\n")
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r" *\n *", "\n", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def display_text_corruption_reasons(text: Any) -> list[str]:
    """Return conservative OCR/display corruption reasons without repairing facts."""
    value = sanitize_display_text(text)
    compact = re.sub(r"\s+", "", value)
    if not compact:
        return ["empty_display_text"]

    reasons: list[str] = []
    upper = value.upper()
    if any(token in upper for token in KNOWN_OCR_NOISE_TOKENS):
        reasons.append("known_ocr_noise_token")
    if re.search(
        r"<\/?[A-Za-z][^>]*>|(?:^|\s)[\[{]\s*[\"'][^\n:]{1,40}[\"']\s*:"
        r"|(?:^|\s)\[\s*(?:[\"'][^\n\]]{1,80}[\"']\s*,\s*)+[\"'][^\n\]]{1,80}[\"']\s*\]",
        value,
    ):
        reasons.append("html_json_or_python_literal")
    if re.search(r"[\ufffd\u0080-\u00a0\u00a2-\u00ff]", value):
        reasons.append("invalid_or_latin1_artifact")
    if re.search(r"[ㄱ-ㅎㅏ-ㅣ]{1,}|(?:\b[가-힣]\s+){3,}[가-힣]\b", value):
        reasons.append("broken_korean_spacing_or_jamo")
    if re.search(r"([^\s])\1{4,}", compact):
        reasons.append("repeated_character")
    words = re.findall(r"\b[^\W_]+\b", value, flags=re.UNICODE)
    if any(count >= 4 for count in Counter(words).values()):
        reasons.append("repeated_token")

    date_patterns = set(
        match.group(0).replace(" ", "")
        for match in re.finditer(
            r"(?:['’`]?\s*(?:19|20)?\d{2})\s*(?:년|[./-])\s*"
            r"\d{1,2}\s*(?:월|[./-])\s*\d{1,2}\s*(?:일)?",
            value,
        )
    )
    month_day_patterns = set(
        match.group(0).replace(" ", "")
        for match in re.finditer(r"(?<!\d)\d{1,2}\s*월\s*\d{1,2}\s*일", value)
    )
    if len(date_patterns) >= 2 or len(month_day_patterns) >= 2:
        reasons.append("multiple_accident_dates")
    heading_count = sum(
        len(re.findall(marker, re.sub(r"\s+", "", value)))
        for marker in ("사고개요", "재해개요", "중대재해발생알림")
    )
    if heading_count >= 2:
        reasons.append("mixed_accident_headings")

    english_tokens = re.findall(r"(?<![A-Za-z0-9])[A-Za-z]{2,12}(?![A-Za-z0-9])", value)
    suspicious_english = []
    for token in english_tokens:
        upper_token = token.upper()
        if upper_token in SAFE_ENGLISH_TOKENS:
            continue
        mixed_case = not (token.islower() or token.isupper() or token.istitle())
        no_vowel = not re.search(r"[AEIOUaeiou]", token)
        short_fragment = len(token) <= 3
        if upper_token in KNOWN_OCR_NOISE_TOKENS or mixed_case or no_vowel or short_fragment:
            suspicious_english.append(token)
    if len(suspicious_english) >= 1 or len(english_tokens) >= 4:
        reasons.append("meaningless_english_fragments")

    isolated_numbers = re.findall(r"(?<!\w)\d{2,}(?!\w)", value)
    symbol_runs = re.findall(r"[^\w\s가-힣.,()%/:;·+\-]{3,}", value)
    if len(isolated_numbers) >= 6 or symbol_runs:
        reasons.append("abnormal_number_or_symbol_sequence")
    noise_chars = re.findall(r"[^A-Za-z0-9가-힣\s.,()%/:;·+\-]", value)
    if compact and len(noise_chars) / len(compact) > 0.06:
        reasons.append("excessive_noise_character_ratio")
    if any(
        marker in re.sub(r"\s+", "", value)
        for marker in ("업종재해유형별", "책자소개", "현장반응", "안전보건주요안내사항")
    ):
        reasons.append("non_case_or_promotional_content")
    return sorted(set(reasons))


def detect_corrupted_ocr_text(text: Any) -> bool:
    return bool(display_text_corruption_reasons(text))


def sanitize_display_case(case: dict[str, Any]) -> dict[str, Any]:
    """Return a copy with display fields normalized, never sourced from raw OCR fields."""
    sanitized = dict(case)
    for field in (
        "display_accident_summary",
        "display_cause_summary",
        "display_prevention_summary",
    ):
        sanitized[field] = sanitize_display_text(case.get(field, ""))
    return sanitized


def _safe_original_image_path(case: dict[str, Any]) -> Path | None:
    relative = str(case.get("original_page_image", "")).strip()
    if not relative:
        return None
    try:
        resolved = (ROOT_DIR / relative).resolve()
        resolved.relative_to(ROOT_DIR.resolve())
    except (OSError, ValueError):
        return None
    return resolved if resolved.is_file() else None


def is_display_safe_case(case: dict[str, Any], require_auto_quality: bool = True) -> bool:
    status = str(case.get("verification_status", "")).strip()
    if status not in PUBLIC_CASE_STATUSES:
        return False
    if not str(case.get("case_id", "")).strip():
        return False
    if not str(case.get("source_document", "")).strip():
        return False
    if case.get("original_page_number") in (None, "") and case.get("page_start") in (None, ""):
        return False
    if _safe_original_image_path(case) is None:
        return False
    sanitized = sanitize_display_case(case)
    summary = sanitized["display_accident_summary"]
    summary_length = len(re.sub(r"\s+", "", summary))
    if not AUTO_SCREENED_MIN_DISPLAY_CHARS <= summary_length <= AUTO_SCREENED_MAX_DISPLAY_CHARS:
        return False
    for field in (
        "display_accident_summary",
        "display_cause_summary",
        "display_prevention_summary",
    ):
        value = sanitized[field]
        if value and detect_corrupted_ocr_text(value):
            return False
    if require_auto_quality and status == AUTO_SCREENED_STATUS:
        try:
            if float(case.get("text_quality_score", -1)) < AUTO_SCREENED_MIN_TEXT_QUALITY:
                return False
            if float(case.get("reading_order_score", -1)) < AUTO_SCREENED_MIN_READING_ORDER:
                return False
            if float(case.get("metadata_quality_score", -1)) < AUTO_SCREENED_MIN_METADATA_QUALITY:
                return False
        except (TypeError, ValueError):
            return False
    return True


def auto_screened_exclusion_reason(
    case: dict[str, Any],
    duplicate_case_ids: set[str] | None = None,
    duplicate_content_hashes: set[str] | None = None,
) -> str:
    status = str(case.get("verification_status", DEFAULT_VERIFICATION_STATUS)).strip()
    if status in {"verified", "manual_review", "rejected"}:
        return f"protected_status_{status}"
    if status not in {DEFAULT_VERIFICATION_STATUS, AUTO_SCREENED_STATUS}:
        return "unsupported_status"
    if str(case.get("mine_relevance", "")) not in {"high", "medium"}:
        return "mine_relevance_not_eligible"
    if str(case.get("ocr_quality_status", "")) != "pass":
        return "ocr_quality_not_pass"
    if bool(case.get("needs_manual_review", True)):
        return "manual_review_required"
    if bool(case.get("needs_reocr", True)):
        return "reocr_required"
    try:
        if float(case.get("text_quality_score", -1)) < AUTO_SCREENED_MIN_TEXT_QUALITY:
            return "low_text_quality"
        if float(case.get("reading_order_score", -1)) < AUTO_SCREENED_MIN_READING_ORDER:
            return "low_reading_order_quality"
        if float(case.get("metadata_quality_score", -1)) < AUTO_SCREENED_MIN_METADATA_QUALITY:
            return "low_metadata_quality"
    except (TypeError, ValueError):
        return "missing_quality_score"

    case_id = str(case.get("case_id", "")).strip()
    content_hash = str(case.get("content_hash", "")).strip()
    if not case_id:
        return "missing_case_id"
    if not content_hash:
        return "missing_content_hash"
    if duplicate_case_ids and case_id in duplicate_case_ids:
        return "duplicate_case_id"
    if duplicate_content_hashes and content_hash in duplicate_content_hashes:
        return "duplicate_content_hash"
    if not str(case.get("source_document", "")).strip():
        return "missing_source_document"
    if case.get("original_page_number") in (None, "") and case.get("page_start") in (None, ""):
        return "missing_source_page"
    if case.get("official_case") is not True:
        return "not_official_case"
    if _safe_original_image_path(case) is None:
        return "missing_original_page_image"

    sanitized = sanitize_display_case(case)
    summary = sanitized["display_accident_summary"]
    summary_length = len(re.sub(r"\s+", "", summary))
    if summary_length < AUTO_SCREENED_MIN_DISPLAY_CHARS:
        return "short_display_summary"
    if summary_length > AUTO_SCREENED_MAX_DISPLAY_CHARS:
        return "long_display_summary"
    industry = re.sub(r"\s+", "", str(case.get("industry", "")))
    if len(industry) <= 1 or industry in {"정보없음", "업종미상", "미상"}:
        return "invalid_industry"
    if bool(case.get("mixed_case_detected", False)):
        return "mixed_case_detected"
    for field in (
        "display_accident_summary",
        "display_cause_summary",
        "display_prevention_summary",
    ):
        value = sanitized[field]
        if field != "display_accident_summary" and not value:
            continue
        reasons = display_text_corruption_reasons(value)
        if reasons:
            return f"corrupted_{field}_{reasons[0]}"
    return ""


def screen_auto_screened_candidates(
    records: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[str, Any], list[dict[str, Any]]]:
    case_id_counts = Counter(str(record.get("case_id", "")).strip() for record in records)
    hash_counts = Counter(str(record.get("content_hash", "")).strip() for record in records)
    duplicate_case_ids = {value for value, count in case_id_counts.items() if value and count > 1}
    duplicate_hashes = {value for value, count in hash_counts.items() if value and count > 1}
    approved: list[dict[str, Any]] = []
    exclusions: Counter[str] = Counter()
    audit_rows: list[dict[str, Any]] = []
    screened_at = _now_iso()
    for record in records:
        reason = auto_screened_exclusion_reason(record, duplicate_case_ids, duplicate_hashes)
        case_id = str(record.get("case_id", "")).strip()
        if reason:
            exclusions[reason] += 1
            audit_rows.append(
                {
                    "case_id": case_id,
                    "screening_status": "excluded",
                    "exclusion_reason": reason,
                    "screened_at": screened_at,
                }
            )
            continue
        item = sanitize_display_case(record)
        item.update(
            {
                "verification_status": AUTO_SCREENED_STATUS,
                "screened_at": screened_at,
                "extraction_quality": "auto_screened_pass",
            }
        )
        if not is_display_safe_case(item):
            exclusions["final_display_safety_check_failed"] += 1
            audit_rows.append(
                {
                    "case_id": case_id,
                    "screening_status": "excluded",
                    "exclusion_reason": "final_display_safety_check_failed",
                    "screened_at": screened_at,
                }
            )
            continue
        approved.append(item)
        audit_rows.append(
            {
                "case_id": case_id,
                "screening_status": AUTO_SCREENED_STATUS,
                "exclusion_reason": "",
                "screened_at": screened_at,
            }
        )
    summary = {
        "input_count": len(records),
        "approved_count": len(approved),
        "excluded_count": len(records) - len(approved),
        "exclusion_reasons": dict(sorted(exclusions.items())),
        "duplicate_case_id_count": len(duplicate_case_ids),
        "duplicate_content_hash_count": len(duplicate_hashes),
    }
    return approved, summary, audit_rows


def write_auto_screening_audit(rows: list[dict[str, Any]]) -> None:
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    temporary_path = AUTO_SCREENING_RESULT_PATH.with_suffix(".jsonl.tmp")
    with temporary_path.open("w", encoding="utf-8", newline="\n") as output:
        for row in rows:
            output.write(json.dumps(row, ensure_ascii=False, sort_keys=True) + "\n")
    temporary_path.replace(AUTO_SCREENING_RESULT_PATH)


def classify_case_relation(
    case: dict[str, Any],
    direct_terms: tuple[str, ...] | list[str],
    analogous_terms: tuple[str, ...] | list[str],
) -> tuple[str, list[str]]:
    searchable = " ".join(
        sanitize_display_text(case.get(field, ""))
        for field in (
            "accident_type",
            "work_type",
            "equipment",
            "display_accident_summary",
            "display_cause_summary",
            "display_prevention_summary",
        )
    ).lower()
    direct_matches = [term for term in direct_terms if term.lower() in searchable]
    if direct_matches:
        return "direct", direct_matches
    analogous_matches = [term for term in analogous_terms if term.lower() in searchable]
    if analogous_matches:
        return "analogous", analogous_matches
    return "", []


def rank_public_official_cases(
    cases: list[dict[str, Any]],
    max_results: int = 3,
) -> list[dict[str, Any]]:
    unique: dict[str, dict[str, Any]] = {}
    for case in cases:
        case_id = str(case.get("case_id", "")).strip()
        tier = PUBLIC_RANK_ORDER.get(
            (str(case.get("relation_type", "")), str(case.get("verification_status", "")))
        )
        if not case_id or tier is None or not is_display_safe_case(case):
            continue
        existing = unique.get(case_id)
        if existing is None:
            unique[case_id] = case
            continue
        existing_tier = PUBLIC_RANK_ORDER.get(
            (
                str(existing.get("relation_type", "")),
                str(existing.get("verification_status", "")),
            ),
            99,
        )
        if tier < existing_tier:
            unique[case_id] = case
    ranked = sorted(
        unique.values(),
        key=lambda item: (
            PUBLIC_RANK_ORDER[
                (str(item.get("relation_type")), str(item.get("verification_status")))
            ],
            -len(item.get("matched_terms", [])),
            float(item["distance"])
            if isinstance(item.get("distance"), (int, float))
            else float("inf"),
            str(item.get("case_id", "")),
        ),
    )
    return ranked[: max(0, min(int(max_results), 3))]


def _find_tesseract() -> Path | None:
    candidates = [
        Path(shutil.which("tesseract") or ""),
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    ]
    return next((path.resolve() for path in candidates if path.is_file()), None)


def _find_pdftoppm() -> Path | None:
    candidates = [
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "native"
        / "poppler"
        / "Library"
        / "bin"
        / "pdftoppm.exe",
        Path(shutil.which("pdftoppm") or ""),
    ]
    return next((path.resolve() for path in candidates if path.is_file()), None)


def _find_tessdata(tesseract_path: Path) -> Path | None:
    candidates = [
        ROOT_DIR / "21_official_accident_case_pipeline" / "tessdata_local",
        tesseract_path.parent / "tessdata",
    ]
    return next(
        (
            path.resolve()
            for path in candidates
            if (path / "kor.traineddata").is_file() and (path / "eng.traineddata").is_file()
        ),
        None,
    )


def _render_page_image(pdf_path: Path, page_number: int, output_path: Path, pdftoppm_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="minesafe_verified_review_") as temp_name:
        safe_pdf = Path(temp_name) / "source.pdf"
        shutil.copyfile(pdf_path, safe_pdf)
        prefix = output_path.with_suffix("")
        subprocess.run(
            [
                str(pdftoppm_path),
                "-f", str(page_number),
                "-l", str(page_number),
                "-singlefile",
                "-png",
                "-r", "300",
                str(safe_pdf),
                str(prefix),
            ],
            check=True,
            capture_output=True,
        )


def _read_tesseract_tsv(image_path: Path, tesseract_path: Path, tessdata_path: Path) -> list[dict[str, Any]]:
    result = subprocess.run(
        [
            str(tesseract_path),
            str(image_path),
            "stdout",
            "--tessdata-dir", str(tessdata_path),
            "-l", "kor+eng",
            "--oem", "1",
            "--psm", "11",
            "tsv",
        ],
        check=True,
        capture_output=True,
    )
    rows: list[dict[str, Any]] = []
    reader = csv.DictReader(io.StringIO(result.stdout.decode("utf-8", errors="replace")), delimiter="\t")
    for row in reader:
        text = str(row.get("text", "")).strip()
        if not text:
            continue
        try:
            confidence = float(row.get("conf", "-1"))
            left = int(row.get("left", "0"))
            top = int(row.get("top", "0"))
            width = int(row.get("width", "0"))
            height = int(row.get("height", "0"))
        except (TypeError, ValueError):
            continue
        if confidence < 0 or width <= 0 or height <= 0:
            continue
        rows.append({"text": text, "left": left, "top": top, "width": width, "height": height})
    return rows


def detect_card_regions(
    rows: list[dict[str, Any]],
    image_width: int,
    image_height: int,
    desired_count: int,
) -> list[tuple[int, int, int, int]]:
    """Create conservative content/card regions from OCR word coordinates."""
    if not rows:
        margin_x = max(1, int(image_width * 0.04))
        margin_y = max(1, int(image_height * 0.04))
        if desired_count <= 1:
            return [(margin_x, margin_y, image_width - margin_x, image_height - margin_y)]
        usable_width = image_width - (margin_x * 2)
        return [
            (
                margin_x + int(usable_width * index / desired_count),
                margin_y,
                margin_x + int(usable_width * (index + 1) / desired_count),
                image_height - margin_y,
            )
            for index in range(desired_count)
        ]

    def bbox(items: list[dict[str, Any]]) -> tuple[int, int, int, int]:
        padding = max(20, int(min(image_width, image_height) * 0.015))
        left = max(0, min(item["left"] for item in items) - padding)
        top = max(0, min(item["top"] for item in items) - padding)
        right = min(image_width, max(item["left"] + item["width"] for item in items) + padding)
        bottom = min(image_height, max(item["top"] + item["height"] for item in items) + padding)
        return left, top, right, bottom

    if desired_count <= 1:
        return [bbox(rows)]
    midpoint = image_width / 2
    left_rows = [row for row in rows if row["left"] + row["width"] / 2 < midpoint]
    right_rows = [row for row in rows if row not in left_rows]
    if left_rows and right_rows:
        regions = [bbox(left_rows), bbox(right_rows)]
    else:
        sorted_rows = sorted(rows, key=lambda row: row["top"])
        split = max(1, len(sorted_rows) // 2)
        regions = [bbox(sorted_rows[:split]), bbox(sorted_rows[split:] or sorted_rows[:split])]
    while len(regions) < desired_count:
        regions.append(regions[-1])
    return regions[:desired_count]


def _text_in_region(rows: list[dict[str, Any]], region: tuple[int, int, int, int]) -> str:
    left, top, right, bottom = region
    selected = [
        row
        for row in rows
        if left <= row["left"] + row["width"] / 2 <= right
        and top <= row["top"] + row["height"] / 2 <= bottom
    ]
    selected.sort(key=lambda row: (row["top"], row["left"]))
    return " ".join(str(row["text"]) for row in selected)


def generate_priority_card_images(records: list[dict[str, Any]]) -> dict[str, Any]:
    """Generate layout-aware source crops for priority cases; status stays unverified."""
    from PIL import Image

    priorities = priority_review_candidates(records)
    tesseract_path = _find_tesseract()
    pdftoppm_path = _find_pdftoppm()
    if not tesseract_path or not pdftoppm_path:
        raise ReviewWorkflowBlocked("Tesseract 또는 PDF 렌더러를 찾지 못했습니다.")
    tessdata_path = _find_tessdata(tesseract_path)
    if not tessdata_path:
        raise ReviewWorkflowBlocked("kor+eng tessdata를 찾지 못했습니다.")

    by_page: dict[tuple[str, int], list[dict[str, Any]]] = {}
    for record in priorities:
        source_file = str(record.get("source_file", "")).strip()
        page_number = int(record.get("original_page_number") or record.get("page_start") or 0)
        if source_file and page_number > 0:
            by_page.setdefault((source_file, page_number), []).append(record)

    generated = 0
    failed: list[str] = []
    CARD_IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    for page_index, ((source_file, page_number), page_records) in enumerate(sorted(by_page.items()), start=1):
        pdf_path = SOURCE_PDF_DIR / source_file
        if not pdf_path.is_file():
            failed.extend(str(record.get("case_id", "")) for record in page_records)
            continue
        with tempfile.TemporaryDirectory(prefix="minesafe_review_page_") as temp_name:
            page_image = Path(temp_name) / f"page_{page_index}.png"
            try:
                _render_page_image(pdf_path, page_number, page_image, pdftoppm_path)
                rows = _read_tesseract_tsv(page_image, tesseract_path, tessdata_path)
                with Image.open(page_image) as image:
                    regions = detect_card_regions(rows, image.width, image.height, len(page_records))
                    for record, region in zip(
                        sorted(page_records, key=lambda item: str(item.get("case_id", ""))),
                        regions,
                    ):
                        case_id = str(record.get("case_id", "")).strip()
                        output_path = CARD_IMAGE_DIR / f"{case_id}_page_{page_number}.png"
                        image.crop(region).save(output_path)
                        relative_image = output_path.relative_to(ROOT_DIR).as_posix()
                        updated = dict(record)
                        updated.update(
                            {
                                "original_page_image": relative_image,
                                "original_page_number": page_number,
                                "layout_ocr_text": _text_in_region(rows, region),
                                "extraction_engine": "tesseract_tsv_layout_unverified",
                                "extraction_quality": "manual_review_required",
                                "review_updated_at": _now_iso(),
                            }
                        )
                        _append_review_event(updated, "layout_card_image_generated")
                        generated += 1
            except Exception:
                failed.extend(str(record.get("case_id", "")) for record in page_records)
    return {
        "priority_case_count": len(priorities),
        "generated_image_count": generated,
        "failed_case_ids": failed,
        "engine": "tesseract_tsv_layout_unverified",
    }


def save_review_update(
    case_id: str,
    edited_fields: dict[str, Any],
    verification_status: str,
    verification_note: str = "",
    verified_fields: list[str] | None = None,
    rejection_reason: str = "",
) -> dict[str, Any]:
    if verification_status not in CASE_STATUSES:
        raise ReviewWorkflowBlocked("허용되지 않은 검증 상태입니다.")
    records = {str(record.get("case_id", "")): record for record in load_review_records()}
    if case_id not in records:
        raise ReviewWorkflowBlocked("검수 대상 case_id를 찾지 못했습니다.")
    record = dict(records[case_id])
    for field in EDITABLE_FIELDS:
        if field in edited_fields:
            record[field] = str(edited_fields[field] or "").strip()
    checks = sorted(set(verified_fields or []))
    if verification_status == "verified":
        if set(checks) != set(REQUIRED_VERIFICATION_CHECKS):
            raise ReviewWorkflowBlocked("원본 대조 필수 확인을 모두 완료해야 합니다.")
        image_path = ROOT_DIR / str(record.get("original_page_image", ""))
        if not image_path.is_file():
            raise ReviewWorkflowBlocked("원본 사고 카드 이미지가 없어 검증 완료할 수 없습니다.")
        for field in ("accident_summary", "source_document", "original_page_number", "case_id"):
            if not record.get(field):
                raise ReviewWorkflowBlocked(f"검증 필수 필드 누락: {field}")
        record["verified_at"] = _now_iso()
        record["display_accident_summary"] = sanitize_display_text(
            record.get("accident_summary", "")
        )
        record["display_cause_summary"] = sanitize_display_text(
            record.get("cause_summary", "")
        ) or "공식 원문에서 별도 원인을 확인하지 못했습니다."
        record["display_prevention_summary"] = sanitize_display_text(
            record.get("prevention_summary", "")
        ) or "공식 원문에서 별도 예방사항을 확인하지 못했습니다."
    elif verification_status != "verified":
        record["verified_at"] = ""
    record.update(
        {
            "verification_status": verification_status,
            "verification_note": str(verification_note or "").strip(),
            "verified_fields": checks,
            "rejection_reason": str(rejection_reason or "").strip(),
            "review_updated_at": _now_iso(),
        }
    )
    _append_review_event(record, "manual_review_update")
    return record


def verified_records_only(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    verified: list[dict[str, Any]] = []
    seen: set[str] = set()
    for record in records:
        case_id = str(record.get("case_id", "")).strip()
        if record.get("verification_status") != "verified" or not case_id or case_id in seen:
            continue
        image_path = ROOT_DIR / str(record.get("original_page_image", ""))
        if not image_path.is_file():
            continue
        if set(record.get("verified_fields", [])) != set(REQUIRED_VERIFICATION_CHECKS):
            continue
        if not str(record.get("display_accident_summary", "")).strip():
            continue
        if not is_display_safe_case(record, require_auto_quality=False):
            continue
        seen.add(case_id)
        verified.append(record)
    return verified


def _verified_document(record: dict[str, Any]) -> str:
    fields = (
        ("사고 유형", record.get("accident_type", "")),
        ("발생일", record.get("accident_date", "")),
        ("업종", record.get("industry", "")),
        ("사고 개요", record.get("display_accident_summary", "")),
        ("발생 원인", record.get("display_cause_summary", "")),
        ("예방사항", record.get("display_prevention_summary", "")),
        ("출처", record.get("source_document", "")),
        ("페이지", record.get("original_page_number", "")),
        ("case_id", record.get("case_id", "")),
    )
    return "\n".join(f"{label}: {str(value or '').strip()}" for label, value in fields)


def rebuild_verified_case_db(records: list[dict[str, Any]]) -> dict[str, Any]:
    """Build a separate DB from explicitly verified records only."""
    verified = verified_records_only(records)
    if not verified:
        if VERIFIED_CASE_DB_DIR.exists():
            backup = VERIFIED_CASE_DB_DIR.with_name(
                f"23_verified_official_accident_case_vector_db_backup_{_timestamp()}"
            )
            VERIFIED_CASE_DB_DIR.rename(backup)
            return {"status": "zero_verified_db_archived", "verified_count": 0, "backup": str(backup)}
        return {"status": "skipped_zero_verified", "verified_count": 0, "db_created": False}
    if VERIFIED_CASE_COLLECTION_NAME in {
        LAW_COLLECTION_NAME,
        SOURCE_CASE_COLLECTION_NAME,
        AUTO_SCREENED_CASE_COLLECTION_NAME,
    }:
        raise ReviewWorkflowBlocked("verified collection 이름이 기존 collection과 충돌합니다.")
    if VERIFIED_CASE_DB_CANDIDATE_DIR.exists():
        raise ReviewWorkflowBlocked("verified candidate DB가 이미 있어 덮어쓰지 않습니다.")

    import chromadb
    from sentence_transformers import SentenceTransformer

    model = SentenceTransformer(EMBEDDING_MODEL_NAME, local_files_only=True)
    VERIFIED_CASE_DB_CANDIDATE_DIR.mkdir(parents=True, exist_ok=False)
    client = chromadb.PersistentClient(path=str(VERIFIED_CASE_DB_CANDIDATE_DIR))
    collection = client.get_or_create_collection(
        name=VERIFIED_CASE_COLLECTION_NAME,
        metadata={"hnsw:space": "cosine", "verification_required": True},
    )
    documents = [_verified_document(record) for record in verified]
    embeddings = model.encode(documents, normalize_embeddings=True, show_progress_bar=False).tolist()
    metadata_fields = (
        "case_id", "official_case", "ocr_quality_status", "source_document",
        "source_file", "source_page_url", "source_period", "page_start", "page_end",
        "original_page_number", "original_page_image", "accident_date", "accident_year",
        "industry", "accident_type", "accident_summary", "cause_summary",
        "prevention_summary", "mine_relevance", "verification_status", "verified_at",
        "verification_note", "verified_fields", "extraction_engine", "extraction_quality",
        "display_accident_summary", "display_cause_summary", "display_prevention_summary",
    )
    collection.add(
        ids=[str(record["case_id"]) for record in verified],
        embeddings=embeddings,
        documents=documents,
        metadatas=[
            {field: _safe_metadata_value(record.get(field)) for field in metadata_fields}
            for record in verified
        ],
    )
    if collection.count() != len(verified):
        raise ReviewWorkflowBlocked("verified candidate DB 건수가 검증 사례 수와 다릅니다.")
    payload = collection.get(include=["metadatas"])
    if any(metadata.get("verification_status") != "verified" for metadata in payload.get("metadatas", [])):
        raise ReviewWorkflowBlocked("미검증 사례가 verified candidate DB에 포함됐습니다.")
    del collection, client
    gc.collect()

    backup_path = ""
    if VERIFIED_CASE_DB_DIR.exists():
        backup = VERIFIED_CASE_DB_DIR.with_name(
            f"23_verified_official_accident_case_vector_db_backup_{_timestamp()}"
        )
        VERIFIED_CASE_DB_DIR.rename(backup)
        backup_path = str(backup)
    VERIFIED_CASE_DB_CANDIDATE_DIR.rename(VERIFIED_CASE_DB_DIR)
    return {
        "status": "verified_db_ready",
        "verified_count": len(verified),
        "collection_count": len(verified),
        "db_path": str(VERIFIED_CASE_DB_DIR),
        "backup": backup_path,
    }


def _auto_screened_document(record: dict[str, Any]) -> str:
    fields = (
        ("사고 유형", record.get("accident_type", "")),
        ("발생일", record.get("accident_date", "")),
        ("업종", record.get("industry", "")),
        ("사고 개요", record.get("display_accident_summary", "")),
        ("발생 원인", record.get("display_cause_summary", "")),
        ("예방사항", record.get("display_prevention_summary", "")),
        ("출처", record.get("source_document", "")),
        ("페이지", record.get("original_page_number") or record.get("page_start", "")),
        ("case_id", record.get("case_id", "")),
    )
    return "\n".join(f"{label}: {sanitize_display_text(value)}" for label, value in fields)


def rebuild_auto_screened_case_db(records: list[dict[str, Any]]) -> dict[str, Any]:
    """Build a separate collection from strict auto-screening results only."""
    approved, screening_summary, audit_rows = screen_auto_screened_candidates(records)
    write_auto_screening_audit(audit_rows)
    if not approved:
        if AUTO_SCREENED_CASE_DB_DIR.exists():
            backup = AUTO_SCREENED_CASE_DB_DIR.with_name(
                f"23_auto_screened_official_accident_case_vector_db_backup_{_timestamp()}"
            )
            AUTO_SCREENED_CASE_DB_DIR.rename(backup)
            return {
                "status": "zero_auto_screened_db_archived",
                "auto_screened_count": 0,
                "backup": str(backup),
                "screening_summary": screening_summary,
            }
        return {
            "status": "skipped_zero_auto_screened",
            "auto_screened_count": 0,
            "db_created": False,
            "screening_summary": screening_summary,
        }
    collection_names = {
        LAW_COLLECTION_NAME,
        SOURCE_CASE_COLLECTION_NAME,
        VERIFIED_CASE_COLLECTION_NAME,
        AUTO_SCREENED_CASE_COLLECTION_NAME,
    }
    if len(collection_names) != 4:
        raise ReviewWorkflowBlocked("네 공식 collection 이름은 서로 분리되어야 합니다.")
    if AUTO_SCREENED_CASE_DB_CANDIDATE_DIR.exists():
        raise ReviewWorkflowBlocked("auto_screened candidate DB가 이미 있어 덮어쓰지 않습니다.")

    import chromadb
    from sentence_transformers import SentenceTransformer

    model = SentenceTransformer(EMBEDDING_MODEL_NAME, local_files_only=True)
    documents = [_auto_screened_document(record) for record in approved]
    embeddings = model.encode(
        documents,
        normalize_embeddings=True,
        show_progress_bar=False,
    ).tolist()
    AUTO_SCREENED_CASE_DB_CANDIDATE_DIR.mkdir(parents=True, exist_ok=False)
    client = chromadb.PersistentClient(path=str(AUTO_SCREENED_CASE_DB_CANDIDATE_DIR))
    collection = client.get_or_create_collection(
        name=AUTO_SCREENED_CASE_COLLECTION_NAME,
        metadata={"hnsw:space": "cosine", "auto_screened_quality_required": True},
    )
    metadata_fields = (
        "case_id", "content_hash", "official_case", "ocr_quality_status",
        "source_document", "source_file", "source_page_url", "source_period",
        "page_start", "page_end", "original_page_number", "original_page_image",
        "accident_date", "accident_year", "industry", "accident_type", "work_type",
        "equipment", "mine_relevance", "verification_status", "screened_at",
        "extraction_engine", "extraction_quality", "text_quality_score",
        "reading_order_score", "metadata_quality_score", "needs_reocr",
        "needs_manual_review", "display_accident_summary", "display_cause_summary",
        "display_prevention_summary",
    )
    collection.add(
        ids=[str(record["case_id"]) for record in approved],
        embeddings=embeddings,
        documents=documents,
        metadatas=[
            {field: _safe_metadata_value(record.get(field)) for field in metadata_fields}
            for record in approved
        ],
    )
    if collection.count() != len(approved):
        raise ReviewWorkflowBlocked("auto_screened candidate DB 건수가 승인 사례 수와 다릅니다.")
    payload = collection.get(include=["metadatas"])
    metadatas = payload.get("metadatas", []) or []
    ids = [str(value) for value in payload.get("ids", []) or []]
    hashes = [str(metadata.get("content_hash", "")) for metadata in metadatas]
    if len(ids) != len(set(ids)) or len(hashes) != len(set(hashes)):
        raise ReviewWorkflowBlocked("auto_screened candidate DB에 중복 사례가 있습니다.")
    if any(metadata.get("verification_status") != AUTO_SCREENED_STATUS for metadata in metadatas):
        raise ReviewWorkflowBlocked("auto_screened 이외 상태가 candidate DB에 포함됐습니다.")
    if any(not is_display_safe_case(metadata) for metadata in metadatas):
        raise ReviewWorkflowBlocked("문자열 안전검사를 통과하지 못한 candidate 사례가 있습니다.")
    del collection, client
    gc.collect()

    backup_path = ""
    if AUTO_SCREENED_CASE_DB_DIR.exists():
        backup = AUTO_SCREENED_CASE_DB_DIR.with_name(
            f"23_auto_screened_official_accident_case_vector_db_backup_{_timestamp()}"
        )
        AUTO_SCREENED_CASE_DB_DIR.rename(backup)
        backup_path = str(backup)
    AUTO_SCREENED_CASE_DB_CANDIDATE_DIR.rename(AUTO_SCREENED_CASE_DB_DIR)
    current_status = {
        str(record.get("case_id", "")): str(record.get("verification_status", ""))
        for record in records
    }
    for record in approved:
        if current_status.get(str(record.get("case_id", ""))) != AUTO_SCREENED_STATUS:
            _append_review_event(record, "auto_screened_quality_passed")
    return {
        "status": "auto_screened_db_ready",
        "auto_screened_count": len(approved),
        "collection_count": len(approved),
        "db_path": str(AUTO_SCREENED_CASE_DB_DIR),
        "backup": backup_path,
        "screening_summary": screening_summary,
    }
