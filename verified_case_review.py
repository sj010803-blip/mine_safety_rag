"""Manual verification workflow for official accident cases.

This module never rewrites the source OCR JSONL and never marks OCR output as
verified automatically.  A verified state can only be created by an explicit
review action that confirms the source image and required fields.
"""

from __future__ import annotations

import csv
import gc
import hashlib
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
TEXT_SAFE_CASE_DB_DIR = ROOT_DIR / "23_text_safe_official_accident_case_vector_db"
TEXT_SAFE_CASE_DB_CANDIDATE_DIR = (
    ROOT_DIR / "23_text_safe_official_accident_case_vector_db_candidate"
)
REVIEW_DIR = ROOT_DIR / "25_verified_case_review"
REVIEW_STATUS_PATH = REVIEW_DIR / "official_case_review_status.jsonl"
AUTO_SCREENING_RESULT_PATH = REVIEW_DIR / "auto_screened_quality_results.jsonl"
TEXT_SAFE_SCREENING_RESULT_PATH = REVIEW_DIR / "text_safe_quality_results.jsonl"
CARD_IMAGE_DIR = REVIEW_DIR / "card_images"
SOURCE_PDF_DIR = ROOT_DIR / "21_official_accident_case_docs"
ORIGINAL_OCR_JSONL_PATH = (
    ROOT_DIR
    / "21_official_accident_case_ocr"
    / "page_text"
    / "official_siren_ocr_pages.jsonl"
)
SOURCE_MANIFEST_PATH = (
    ROOT_DIR
    / "21_official_accident_case_pipeline"
    / "official_siren_source_manifest.json"
)
PRESENTATION_RECOVERY_PLAN_PATH = REVIEW_DIR / "presentation_case_recovery_plan.json"
PRESENTATION_RECOVERY_AUDIT_PATH = REVIEW_DIR / "presentation_case_recovery_audit.jsonl"
PRESENTATION_RECOVERED_CARD_DIR = CARD_IMAGE_DIR / "presentation_recovered"

LAW_COLLECTION_NAME = "mine_safety_docs"
SOURCE_CASE_COLLECTION_NAME = "mine_official_accident_cases"
VERIFIED_CASE_COLLECTION_NAME = "mine_verified_official_accident_cases"
AUTO_SCREENED_CASE_COLLECTION_NAME = "mine_auto_screened_official_accident_cases"
TEXT_SAFE_CASE_COLLECTION_NAME = "mine_text_safe_official_accident_cases"
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"

VERIFICATION_STATUSES = ("unverified", "verified", "rejected", "manual_review")
AUTO_SCREENED_STATUS = "auto_screened"
CASE_STATUSES = (*VERIFICATION_STATUSES, AUTO_SCREENED_STATUS)
PUBLIC_CASE_STATUSES = ("verified", AUTO_SCREENED_STATUS)
DEFAULT_VERIFICATION_STATUS = "unverified"
VERIFIED_PUBLIC_TIER = "verified"
AUTO_SCREENED_PUBLIC_TIER = AUTO_SCREENED_STATUS
TEXT_SAFE_FALLBACK_TIER = "text_safe_fallback"
HIDDEN_PUBLIC_TIER = "hidden"
PUBLIC_CASE_TIERS = (
    VERIFIED_PUBLIC_TIER,
    AUTO_SCREENED_PUBLIC_TIER,
    TEXT_SAFE_FALLBACK_TIER,
    HIDDEN_PUBLIC_TIER,
)
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
    ("direct", VERIFIED_PUBLIC_TIER): 0,
    ("direct", AUTO_SCREENED_PUBLIC_TIER): 1,
    ("direct", TEXT_SAFE_FALLBACK_TIER): 2,
    ("analogous", VERIFIED_PUBLIC_TIER): 3,
    ("analogous", AUTO_SCREENED_PUBLIC_TIER): 4,
    ("analogous", TEXT_SAFE_FALLBACK_TIER): 5,
    ("broad_family", TEXT_SAFE_FALLBACK_TIER): 6,
}
RISK_FAMILY_TERMS = {
    "mechanical_entanglement": (
        "컨베이어", "벨트", "회전체", "말림", "기계 청소", "방호장치",
        "에너지 차단", "정비 중 재가동", "기계설비",
    ),
    "vehicle_transport": (
        "후진", "신호수", "사각지대", "덤프트럭", "지게차", "굴착기",
        "운반장비", "차량", "충돌", "깔림",
    ),
    "collapse_falling": (
        "낙반", "붕락", "붕괴", "매몰", "토사", "암석", "굴착면",
        "물체 낙하", "무너짐",
    ),
    "electrical_energy": (
        "감전", "누전", "전기설비", "전원 투입", "잠금", "표지",
        "정비 중 재가동",
    ),
    "asphyxiation_gas": (
        "산소결핍", "질식", "밀폐공간", "유해가스", "환기 불량", "가스 중독",
    ),
    "fire_explosion": (
        "발파", "불발공", "화약", "폭발", "화재", "점화원", "폭발물",
    ),
    "fall_from_height": (
        "떨어짐", "추락", "고소작업", "개구부", "사다리", "작업발판",
    ),
}

ELECTRICAL_CASE_ANCHOR_TERMS = (
    "감전", "누전", "전기설비", "전원", "충전부", "전선", "전기",
    "활선", "접지", "절연", "차단기", "케이블", "배선",
)
CONTROLLED_RECOVERY_INDUSTRIES = {
    "건설업", "제조업", "기타업종", "광업", "채석업", "운수업", "서비스업",
}
CONTROLLED_ACCIDENT_TYPE_PATTERNS = {
    "끼임": (r"끼임", r"끼어"),
    "부딪힘": (r"부딪힘", r"부딪혀", r"충돌"),
    "깔림": (r"깔림", r"깔려"),
    "떨어짐": (r"떨어짐", r"떨어져", r"추락"),
    "무너짐": (r"무너짐", r"무너지"),
    "매몰": (r"매몰",),
    "감전": (r"감전",),
    "질식": (r"질식",),
    "폭발": (r"폭발",),
    "화재": (r"화재",),
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
            "public_case_tier": HIDDEN_PUBLIC_TIER,
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
    group_counts = {group_name: 0 for group_name, _keywords, _types in groups}

    # Once the presentation set has been selected and its source images were
    # prepared, keep that exact set stable across later metadata corrections.
    persisted = sorted(
        (
            record
            for record in records
            if str(record.get("priority_review_group", "")) in group_counts
            and str(record.get("case_id", "")).strip()
            and record.get("verification_status") != "rejected"
        ),
        key=lambda record: (
            tuple(group_counts).index(str(record.get("priority_review_group", ""))),
            str(record.get("case_id", "")),
        ),
    )
    for record in persisted:
        case_id = str(record.get("case_id", "")).strip()
        group_name = str(record.get("priority_review_group", ""))
        if case_id in used or group_counts[group_name] >= PRIORITY_GROUP_LIMIT:
            continue
        selected.append(dict(record))
        used.add(case_id)
        group_counts[group_name] += 1

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
        remaining = PRIORITY_GROUP_LIMIT - group_counts[group_name]
        for _score, record in candidates[:remaining]:
            item = dict(record)
            item["priority_review_group"] = group_name
            selected.append(item)
            used.add(str(record.get("case_id", "")))
            group_counts[group_name] += 1
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
    if re.search(r"[\ufffd\u0080-\u00a0\u00a2-\u00b6\u00b8-\u00ff]", value):
        reasons.append("invalid_or_latin1_artifact")
    if re.search(r"[ㄱ-ㅎㅏ-ㅣ]{1,}|(?:\b[가-힣]\s+){3,}[가-힣]\b", value):
        reasons.append("broken_korean_spacing_or_jamo")
    if re.search(r"([^\s])\1{4,}", compact):
        reasons.append("repeated_character")
    words = re.findall(r"\b[^\W_]+\b", value, flags=re.UNICODE)
    if any(count >= 4 for count in Counter(words).values()):
        reasons.append("repeated_token")
    sentence_units = [
        re.sub(r"\s+", "", sentence).strip()
        for sentence in re.split(r"[.!?\n]+", value)
        if len(re.sub(r"\s+", "", sentence)) >= 10
    ]
    if any(count >= 2 for count in Counter(sentence_units).values()):
        reasons.append("repeated_sentence")

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


def effective_public_case_tier(case: dict[str, Any]) -> str:
    """Resolve a public tier without changing the manual verification status."""
    verification_status = str(case.get("verification_status", "")).strip()
    if verification_status in {"manual_review", "rejected"}:
        return HIDDEN_PUBLIC_TIER
    if verification_status == "verified":
        return VERIFIED_PUBLIC_TIER
    if verification_status == AUTO_SCREENED_STATUS:
        return AUTO_SCREENED_PUBLIC_TIER
    explicit_tier = str(case.get("public_case_tier", "")).strip()
    if explicit_tier == TEXT_SAFE_FALLBACK_TIER:
        return TEXT_SAFE_FALLBACK_TIER
    return HIDDEN_PUBLIC_TIER


def text_safe_fallback_exclusion_reason(
    case: dict[str, Any],
    duplicate_case_ids: set[str] | None = None,
    duplicate_content_hashes: set[str] | None = None,
) -> str:
    """Check display text and traceability without requiring optional metadata."""
    verification_status = str(
        case.get("verification_status", DEFAULT_VERIFICATION_STATUS)
    ).strip()
    if verification_status in {"manual_review", "rejected"}:
        return f"protected_status_{verification_status}"
    if verification_status in {"verified", AUTO_SCREENED_STATUS}:
        return "higher_public_tier"
    if verification_status != DEFAULT_VERIFICATION_STATUS:
        return "unsupported_verification_status"
    if case.get("official_case") is not True:
        return "not_official_case"

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
    if str(case.get("ocr_quality_status", "")) != "pass":
        return "ocr_quality_not_pass"
    if bool(case.get("mixed_case_detected", False)):
        return "mixed_case_detected"

    sanitized = sanitize_display_case(case)
    summary = sanitized["display_accident_summary"]
    summary_length = len(re.sub(r"\s+", "", summary))
    if summary_length < AUTO_SCREENED_MIN_DISPLAY_CHARS:
        return "short_display_summary"
    if summary_length > AUTO_SCREENED_MAX_DISPLAY_CHARS:
        return "long_display_summary"
    for field in (
        "display_accident_summary",
        "display_cause_summary",
        "display_prevention_summary",
    ):
        value = sanitized[field]
        if not value and field != "display_accident_summary":
            continue
        reasons = display_text_corruption_reasons(value)
        if reasons:
            return f"corrupted_{field}_{reasons[0]}"
    return ""


def screen_text_safe_fallback_candidates(
    records: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[str, Any], list[dict[str, Any]]]:
    """Select only traceable, display-safe fallback records; statuses stay unchanged."""
    case_id_counts = Counter(str(record.get("case_id", "")).strip() for record in records)
    hash_counts = Counter(str(record.get("content_hash", "")).strip() for record in records)
    duplicate_case_ids = {value for value, count in case_id_counts.items() if value and count > 1}
    duplicate_hashes = {value for value, count in hash_counts.items() if value and count > 1}
    approved: list[dict[str, Any]] = []
    exclusions: Counter[str] = Counter()
    audit_rows: list[dict[str, Any]] = []
    screened_at = _now_iso()
    for record in records:
        reason = text_safe_fallback_exclusion_reason(
            record,
            duplicate_case_ids,
            duplicate_hashes,
        )
        case_id = str(record.get("case_id", "")).strip()
        if reason:
            exclusions[reason] += 1
            audit_rows.append(
                {
                    "case_id": case_id,
                    "public_case_tier": HIDDEN_PUBLIC_TIER,
                    "exclusion_reason": reason,
                    "screened_at": screened_at,
                }
            )
            continue
        item = sanitize_display_case(record)
        item.update(
            {
                "public_case_tier": TEXT_SAFE_FALLBACK_TIER,
                "text_safe_screened_at": screened_at,
            }
        )
        if not is_public_display_safe_case(item):
            exclusions["final_display_safety_check_failed"] += 1
            audit_rows.append(
                {
                    "case_id": case_id,
                    "public_case_tier": HIDDEN_PUBLIC_TIER,
                    "exclusion_reason": "final_display_safety_check_failed",
                    "screened_at": screened_at,
                }
            )
            continue
        approved.append(item)
        audit_rows.append(
            {
                "case_id": case_id,
                "public_case_tier": TEXT_SAFE_FALLBACK_TIER,
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


def write_text_safe_screening_audit(rows: list[dict[str, Any]]) -> None:
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    temporary_path = TEXT_SAFE_SCREENING_RESULT_PATH.with_suffix(".jsonl.tmp")
    with temporary_path.open("w", encoding="utf-8", newline="\n") as output:
        for row in rows:
            output.write(json.dumps(row, ensure_ascii=False, sort_keys=True) + "\n")
    temporary_path.replace(TEXT_SAFE_SCREENING_RESULT_PATH)


def is_public_display_safe_case(case: dict[str, Any]) -> bool:
    tier = effective_public_case_tier(case)
    if tier in {VERIFIED_PUBLIC_TIER, AUTO_SCREENED_PUBLIC_TIER}:
        return is_display_safe_case(case)
    if tier == TEXT_SAFE_FALLBACK_TIER:
        return text_safe_fallback_exclusion_reason(case) == ""
    return False


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
    direct_term_set = {term.lower() for term in direct_terms}
    if "컨베이어" in direct_term_set and direct_matches:
        conveyor_specific = (
            "컨베이어", "벨트", "말림", "회전체", "정비 중 재가동",
            "청소 중 기계 작동", "방호장치", "에너지 차단", "전원 미차단",
        )
        if not any(term in searchable for term in conveyor_specific):
            direct_matches = []
    if "후진" in direct_term_set and direct_matches:
        vehicle_motion_terms = (
            "후진", "신호수", "덤프트럭", "지게차", "차량 충돌",
            "사각지대", "작업자 충돌",
        )
        excavator_motion = "굴착기" in searchable and any(
            term in searchable for term in ("후진", "충돌", "부딪힘", "깔림", "사각지대")
        )
        if not any(term in searchable for term in vehicle_motion_terms) and not excavator_motion:
            direct_matches = []
    if direct_matches:
        return "direct", direct_matches
    analogous_matches = [term for term in analogous_terms if term.lower() in searchable]
    if analogous_matches:
        return "analogous", analogous_matches
    return "", []


def classify_case_risk_families(case: dict[str, Any]) -> set[str]:
    """Classify only explicit risk mechanisms found in public display fields."""
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
    families: set[str] = set()
    for family, terms in RISK_FAMILY_TERMS.items():
        matches = [term for term in terms if term.lower() in searchable]
        if not matches:
            continue
        if family == "vehicle_transport" and not any(
            term in searchable
            for term in (
                "후진", "신호수", "사각지대", "덤프트럭", "지게차",
                "굴착기", "운반장비", "차량", "중장비",
            )
        ):
            continue
        if family == "electrical_energy" and not any(
            term in searchable for term in ELECTRICAL_CASE_ANCHOR_TERMS
        ):
            continue
        families.add(family)
    return families


def classify_public_case_relation(
    case: dict[str, Any],
    direct_terms: tuple[str, ...] | list[str],
    analogous_terms: tuple[str, ...] | list[str],
    question_risk_family: str = "",
) -> tuple[str, list[str]]:
    case_risk_families = (
        classify_case_risk_families(case) if question_risk_family else set()
    )
    relation_type, matched_terms = classify_case_relation(case, direct_terms, analogous_terms)
    if relation_type:
        # "잠금"·"표지" 같은 일반 안전용어만 일치한 비전기 사례가
        # 전기안전 유사 사례로 노출되지 않도록 실제 전기 문맥을 재확인한다.
        if (
            question_risk_family == "electrical_energy"
            and question_risk_family not in case_risk_families
        ):
            return "", []
        return relation_type, matched_terms
    if (
        effective_public_case_tier(case) == TEXT_SAFE_FALLBACK_TIER
        and question_risk_family
        and question_risk_family in case_risk_families
    ):
        searchable = " ".join(
            sanitize_display_text(case.get(field, "")).lower()
            for field in (
                "accident_type",
                "work_type",
                "equipment",
                "display_accident_summary",
                "display_cause_summary",
                "display_prevention_summary",
            )
        )
        family_matches = [
            term
            for term in RISK_FAMILY_TERMS.get(question_risk_family, ())
            if term.lower() in searchable
        ]
        return "broad_family", family_matches
    return "", []


def rank_public_official_cases(
    cases: list[dict[str, Any]],
    max_results: int = 3,
) -> list[dict[str, Any]]:
    unique: dict[str, dict[str, Any]] = {}
    for case in cases:
        case_id = str(case.get("case_id", "")).strip()
        public_tier = effective_public_case_tier(case)
        tier = PUBLIC_RANK_ORDER.get(
            (str(case.get("relation_type", "")), public_tier)
        )
        if not case_id or tier is None or not is_public_display_safe_case(case):
            continue
        existing = unique.get(case_id)
        if existing is None:
            unique[case_id] = case
            continue
        existing_tier = PUBLIC_RANK_ORDER.get(
            (
                str(existing.get("relation_type", "")),
                effective_public_case_tier(existing),
            ),
            99,
        )
        if tier < existing_tier:
            unique[case_id] = case
    ranked = sorted(
        unique.values(),
        key=lambda item: (
            PUBLIC_RANK_ORDER[
                (str(item.get("relation_type")), effective_public_case_tier(item))
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


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as input_file:
        for chunk in iter(lambda: input_file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_official_source_manifest_index(
    manifest_path: Path = SOURCE_MANIFEST_PATH,
) -> dict[str, dict[str, Any]]:
    """Load only the exact official manifest and index it by saved filename."""
    payload = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    if not isinstance(payload, list):
        raise ReviewWorkflowBlocked("공식 출처 manifest 형식이 올바르지 않습니다.")
    index: dict[str, dict[str, Any]] = {}
    for row in payload:
        if not isinstance(row, dict):
            continue
        filename = str(row.get("saved_filename", "")).strip()
        if not filename or filename in index:
            raise ReviewWorkflowBlocked("공식 출처 manifest 파일명이 없거나 중복됐습니다.")
        index[filename] = dict(row)
    return index


def official_industry_from_manifest(
    source_file: str,
    manifest_entry: dict[str, Any],
) -> str:
    """Infer industry only when an exact category-specific source file proves it."""
    if str(manifest_entry.get("saved_filename", "")).strip() != str(source_file).strip():
        return ""
    source_id = str(manifest_entry.get("source_id", "")).upper()
    suffix_map = {
        "-CONSTRUCTION": "건설업",
        "-MANUFACTURING": "제조업",
        "-OTHER": "기타업종",
    }
    return next((industry for suffix, industry in suffix_map.items() if source_id.endswith(suffix)), "")


def source_period_from_manifest(manifest_entry: dict[str, Any]) -> str:
    start = str(manifest_entry.get("coverage_start", "")).strip()
    end = str(manifest_entry.get("coverage_end", "")).strip()
    if re.fullmatch(r"\d{4}-01-01", start) and end == f"{start[:4]}-12-31":
        return start[:4]
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", start) and end.startswith(start[:7]):
        return start[:7]
    return str(manifest_entry.get("source_period", "")).strip() or start[:7]


def explicit_accident_type_for_recovery(text: str, requested_type: str = "") -> str:
    """Return a controlled type only when its source wording is explicit in display text."""
    value = sanitize_display_text(text)
    matches = {
        accident_type
        for accident_type, patterns in CONTROLLED_ACCIDENT_TYPE_PATTERNS.items()
        if any(re.search(pattern, value) for pattern in patterns)
    }
    requested = str(requested_type).strip()
    if requested:
        return requested if requested in matches else ""
    return next(iter(matches)) if len(matches) == 1 else ""


def load_presentation_recovery_plan(
    plan_path: Path = PRESENTATION_RECOVERY_PLAN_PATH,
) -> list[dict[str, Any]]:
    payload = json.loads(plan_path.read_text(encoding="utf-8-sig"))
    cases = payload.get("cases", []) if isinstance(payload, dict) else []
    if not isinstance(cases, list) or len(cases) > PRIORITY_GROUP_LIMIT * 4:
        raise ReviewWorkflowBlocked("발표 우선 복구 계획은 최대 12건이어야 합니다.")
    case_ids = [str(case.get("case_id", "")).strip() for case in cases if isinstance(case, dict)]
    if not case_ids or any(not case_id for case_id in case_ids) or len(case_ids) != len(set(case_ids)):
        raise ReviewWorkflowBlocked("발표 우선 복구 계획의 case_id가 없거나 중복됐습니다.")
    return [dict(case) for case in cases if isinstance(case, dict)]


def _normalized_crop_box(value: Any) -> tuple[float, float, float, float] | None:
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        return None
    try:
        left, top, right, bottom = (float(part) for part in value)
    except (TypeError, ValueError):
        return None
    if not (0 <= left < right <= 1 and 0 <= top < bottom <= 1):
        return None
    if (right - left) < 0.25 or (bottom - top) < 0.5:
        return None
    return left, top, right, bottom


def _summary_contains_iso_date(summary: str, iso_date: str) -> bool:
    match = re.fullmatch(r"(20\d{2})-(\d{2})-(\d{2})", str(iso_date).strip())
    if not match:
        return False
    year, month, day = (int(part) for part in match.groups())
    compact = re.sub(r"\s+", "", summary)
    return f"{year}년{month}월{day}일" in compact


def write_presentation_recovery_audit(rows: list[dict[str, Any]]) -> None:
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    temporary = PRESENTATION_RECOVERY_AUDIT_PATH.with_suffix(".jsonl.tmp")
    with temporary.open("w", encoding="utf-8", newline="\n") as output:
        for row in rows:
            output.write(json.dumps(row, ensure_ascii=False, sort_keys=True) + "\n")
    temporary.replace(PRESENTATION_RECOVERY_AUDIT_PATH)


def recover_presentation_priority_cases(
    records: list[dict[str, Any]],
    recovery_specs: list[dict[str, Any]],
    *,
    manifest_index: dict[str, dict[str, Any]] | None = None,
    output_dir: Path = PRESENTATION_RECOVERED_CARD_DIR,
    pdftoppm_path: Path | None = None,
    persist_events: bool = True,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Recover only explicitly listed priority cases from exact official card images."""
    if len(recovery_specs) > PRIORITY_GROUP_LIMIT * 4:
        raise ReviewWorkflowBlocked("발표 우선 복구 대상은 최대 12건입니다.")
    manifest_index = manifest_index or load_official_source_manifest_index()
    priority_ids = {
        str(record.get("case_id", "")).strip()
        for record in priority_review_candidates(records)
    }
    by_id = {str(record.get("case_id", "")).strip(): dict(record) for record in records}
    spec_ids = [str(spec.get("case_id", "")).strip() for spec in recovery_specs]
    if any(not case_id for case_id in spec_ids) or len(spec_ids) != len(set(spec_ids)):
        raise ReviewWorkflowBlocked("발표 우선 복구 대상 case_id가 없거나 중복됐습니다.")
    renderer = pdftoppm_path or _find_pdftoppm()
    if not renderer:
        raise ReviewWorkflowBlocked("공식 PDF 페이지 렌더러를 찾지 못했습니다.")

    from PIL import Image

    successes: list[str] = []
    audit_rows: list[dict[str, Any]] = []
    rendered_pages: dict[tuple[str, int], Path] = {}
    verified_pdf_hashes: set[str] = set()
    output_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="minesafe_presentation_recovery_") as temp_name:
        temporary_dir = Path(temp_name)
        for spec in recovery_specs:
            case_id = str(spec.get("case_id", "")).strip()
            reason = ""
            record = by_id.get(case_id)
            if case_id not in priority_ids:
                reason = "not_presentation_priority_case"
            elif record is None:
                reason = "case_id_not_found"
            elif str(record.get("verification_status", DEFAULT_VERIFICATION_STATUS)) in {
                "verified", "manual_review", "rejected"
            }:
                reason = "protected_verification_status"

            source_file = str(spec.get("source_file", "")).strip()
            try:
                page_number = int(spec.get("page_start", 0))
            except (TypeError, ValueError):
                page_number = 0
            manifest_entry = manifest_index.get(source_file, {})
            pdf_path = SOURCE_PDF_DIR / source_file
            crop_box = _normalized_crop_box(spec.get("crop_box"))
            if not reason and (
                source_file != str(record.get("source_file", "")).strip()
                or page_number != int(record.get("page_start") or 0)
            ):
                reason = "source_or_page_mismatch"
            if not reason and not manifest_entry:
                reason = "manifest_source_not_found"
            if not reason and not (1 <= page_number <= int(manifest_entry.get("page_count", 0))):
                reason = "page_out_of_manifest_range"
            if not reason and (not pdf_path.is_file() or not crop_box):
                reason = "source_pdf_or_crop_missing"
            if not reason and source_file not in verified_pdf_hashes:
                expected_hash = str(manifest_entry.get("sha256", "")).lower()
                if not expected_hash or _sha256_file(pdf_path).lower() != expected_hash:
                    reason = "source_pdf_hash_mismatch"
                else:
                    verified_pdf_hashes.add(source_file)

            summary = sanitize_display_text(spec.get("display_accident_summary", ""))
            summary_length = len(re.sub(r"\s+", "", summary))
            corruption_reasons = display_text_corruption_reasons(summary)
            if not reason and not AUTO_SCREENED_MIN_DISPLAY_CHARS <= summary_length <= AUTO_SCREENED_MAX_DISPLAY_CHARS:
                reason = "display_summary_length_out_of_range"
            if not reason and corruption_reasons:
                reason = f"corrupted_display_summary_{corruption_reasons[0]}"

            manifest_industry = official_industry_from_manifest(source_file, manifest_entry)
            planned_industry = str(spec.get("industry", "")).strip()
            if manifest_industry:
                industry = manifest_industry
                if planned_industry and planned_industry != manifest_industry:
                    reason = reason or "industry_conflicts_with_manifest"
            elif str(spec.get("industry_source", "")) == "original_card_image":
                industry = planned_industry
            else:
                industry = ""
            if not reason and industry not in CONTROLLED_RECOVERY_INDUSTRIES:
                reason = "industry_not_officially_confirmed"

            accident_date = str(spec.get("accident_date", "")).strip()
            accident_type = explicit_accident_type_for_recovery(
                summary,
                str(spec.get("accident_type", "")).strip(),
            )
            if not reason and not _summary_contains_iso_date(summary, accident_date):
                reason = "accident_date_not_explicit_in_summary"
            if not reason and not accident_type:
                reason = "accident_type_not_explicit_in_summary"

            if reason:
                audit_rows.append({"case_id": case_id, "status": "excluded", "reason": reason})
                continue

            page_key = (source_file, page_number)
            if page_key not in rendered_pages:
                page_image = temporary_dir / f"page_{len(rendered_pages) + 1}.png"
                _render_page_image(pdf_path, page_number, page_image, renderer)
                rendered_pages[page_key] = page_image
            page_image = rendered_pages[page_key]
            with Image.open(page_image) as image:
                left, top, right, bottom = crop_box
                pixel_box = (
                    round(image.width * left),
                    round(image.height * top),
                    round(image.width * right),
                    round(image.height * bottom),
                )
                card = image.crop(pixel_box)
                if card.width < 600 or card.height < 900 or card.getbbox() is None:
                    audit_rows.append({"case_id": case_id, "status": "excluded", "reason": "invalid_card_crop"})
                    continue
                output_path = output_dir / f"{case_id}_page_{page_number}.png"
                temporary_output = temporary_dir / output_path.name
                card.save(temporary_output)
                shutil.copyfile(temporary_output, output_path)

            updated = dict(record)
            updated.update(
                {
                    "source_document": str(manifest_entry.get("title", "")).strip(),
                    "source_period": source_period_from_manifest(manifest_entry),
                    "source_file": source_file,
                    "source_page_url": str(manifest_entry.get("source_page_url", "")).strip(),
                    "publisher": str(manifest_entry.get("publisher", "")).strip(),
                    "page_start": page_number,
                    "page_end": page_number,
                    "original_page_number": page_number,
                    "original_page_image": output_path.relative_to(ROOT_DIR).as_posix(),
                    "industry": industry,
                    "accident_date": accident_date,
                    "accident_year": int(accident_date[:4]),
                    "accident_month": int(accident_date[5:7]),
                    "accident_type": accident_type,
                    "display_accident_summary": summary,
                    "display_cause_summary": "",
                    "display_prevention_summary": "",
                    "text_quality_score": 100,
                    "reading_order_score": 100,
                    "metadata_quality_score": 100,
                    "needs_reocr": False,
                    "needs_manual_review": False,
                    "mixed_case_detected": False,
                    "quality_reasons": [],
                    "extraction_engine": "official_pdf_card_crop_exact_transcription",
                    "extraction_quality": "presentation_source_recovered_unverified",
                    "review_updated_at": _now_iso(),
                }
            )
            if auto_screened_exclusion_reason(updated) not in {"", "missing_original_page_image"}:
                audit_rows.append(
                    {
                        "case_id": case_id,
                        "status": "excluded",
                        "reason": auto_screened_exclusion_reason(updated),
                    }
                )
                output_path.unlink(missing_ok=True)
                continue
            by_id[case_id] = updated
            successes.append(case_id)
            audit_rows.append({"case_id": case_id, "status": "recovered_unverified", "reason": ""})
            if persist_events:
                _append_review_event(updated, "presentation_source_recovered_unverified")

    write_presentation_recovery_audit(audit_rows)
    return (
        sorted(by_id.values(), key=lambda item: str(item.get("case_id", ""))),
        {
            "priority_case_count": len(priority_ids),
            "planned_case_count": len(recovery_specs),
            "recovered_case_count": len(successes),
            "recovered_case_ids": successes,
            "excluded_count": len(recovery_specs) - len(successes),
            "audit_rows": audit_rows,
        },
    )


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
        TEXT_SAFE_CASE_COLLECTION_NAME,
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
    del collection
    client.close()
    del client
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
        TEXT_SAFE_CASE_COLLECTION_NAME,
    }
    if len(collection_names) != 5:
        raise ReviewWorkflowBlocked("다섯 공식 collection 이름은 서로 분리되어야 합니다.")
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
    del collection
    client.close()
    del client
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


def rebuild_text_safe_case_db(records: list[dict[str, Any]]) -> dict[str, Any]:
    """Build a separate DB from display-safe fallback cases without changing statuses."""
    approved, screening_summary, audit_rows = screen_text_safe_fallback_candidates(records)
    write_text_safe_screening_audit(audit_rows)
    if not approved:
        if TEXT_SAFE_CASE_DB_DIR.exists():
            backup = TEXT_SAFE_CASE_DB_DIR.with_name(
                f"23_text_safe_official_accident_case_vector_db_backup_{_timestamp()}"
            )
            TEXT_SAFE_CASE_DB_DIR.rename(backup)
            return {
                "status": "zero_text_safe_db_archived",
                "text_safe_count": 0,
                "backup": str(backup),
                "screening_summary": screening_summary,
            }
        return {
            "status": "skipped_zero_text_safe",
            "text_safe_count": 0,
            "db_created": False,
            "screening_summary": screening_summary,
        }

    collection_names = {
        LAW_COLLECTION_NAME,
        SOURCE_CASE_COLLECTION_NAME,
        VERIFIED_CASE_COLLECTION_NAME,
        AUTO_SCREENED_CASE_COLLECTION_NAME,
        TEXT_SAFE_CASE_COLLECTION_NAME,
    }
    if len(collection_names) != 5:
        raise ReviewWorkflowBlocked("다섯 공식 collection 이름은 서로 분리되어야 합니다.")
    if TEXT_SAFE_CASE_DB_CANDIDATE_DIR.exists():
        raise ReviewWorkflowBlocked("text_safe candidate DB가 이미 있어 덮어쓰지 않습니다.")

    import chromadb
    from sentence_transformers import SentenceTransformer

    model = SentenceTransformer(EMBEDDING_MODEL_NAME, local_files_only=True)
    documents = [_auto_screened_document(record) for record in approved]
    embeddings = model.encode(
        documents,
        normalize_embeddings=True,
        show_progress_bar=False,
    ).tolist()
    TEXT_SAFE_CASE_DB_CANDIDATE_DIR.mkdir(parents=True, exist_ok=False)
    client = chromadb.PersistentClient(path=str(TEXT_SAFE_CASE_DB_CANDIDATE_DIR))
    collection = client.get_or_create_collection(
        name=TEXT_SAFE_CASE_COLLECTION_NAME,
        metadata={"hnsw:space": "cosine", "display_text_safety_required": True},
    )
    metadata_fields = (
        "case_id", "content_hash", "official_case", "ocr_quality_status",
        "source_document", "source_file", "source_page_url", "source_period",
        "page_start", "page_end", "original_page_number", "original_page_image",
        "accident_date", "accident_year", "industry", "industry_detail",
        "accident_type", "work_type", "equipment", "mine_relevance",
        "verification_status", "public_case_tier", "text_safe_screened_at",
        "display_accident_summary", "display_cause_summary",
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
        raise ReviewWorkflowBlocked("text_safe candidate DB 건수가 안전 사례 수와 다릅니다.")
    payload = collection.get(include=["metadatas"])
    metadatas = payload.get("metadatas", []) or []
    ids = [str(value) for value in payload.get("ids", []) or []]
    hashes = [str(metadata.get("content_hash", "")) for metadata in metadatas]
    if len(ids) != len(set(ids)) or len(hashes) != len(set(hashes)):
        raise ReviewWorkflowBlocked("text_safe candidate DB에 중복 사례가 있습니다.")
    if any(effective_public_case_tier(metadata) != TEXT_SAFE_FALLBACK_TIER for metadata in metadatas):
        raise ReviewWorkflowBlocked("text_safe 이외 등급이 candidate DB에 포함됐습니다.")
    if any(not is_public_display_safe_case(metadata) for metadata in metadatas):
        raise ReviewWorkflowBlocked("text_safe candidate DB에 깨진 문자열이 포함됐습니다.")
    del collection
    client.close()
    del client
    gc.collect()

    backup_path = ""
    if TEXT_SAFE_CASE_DB_DIR.exists():
        backup = TEXT_SAFE_CASE_DB_DIR.with_name(
            f"23_text_safe_official_accident_case_vector_db_backup_{_timestamp()}"
        )
        TEXT_SAFE_CASE_DB_DIR.rename(backup)
        backup_path = str(backup)
    TEXT_SAFE_CASE_DB_CANDIDATE_DIR.rename(TEXT_SAFE_CASE_DB_DIR)
    return {
        "status": "text_safe_db_ready",
        "text_safe_count": len(approved),
        "collection_count": len(approved),
        "db_path": str(TEXT_SAFE_CASE_DB_DIR),
        "backup": backup_path,
        "screening_summary": screening_summary,
    }
