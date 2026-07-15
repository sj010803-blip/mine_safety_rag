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
REVIEW_DIR = ROOT_DIR / "25_verified_case_review"
REVIEW_STATUS_PATH = REVIEW_DIR / "official_case_review_status.jsonl"
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
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"

VERIFICATION_STATUSES = ("unverified", "verified", "rejected", "manual_review")
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
    if verification_status not in VERIFICATION_STATUSES:
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
        if not str(record.get("accident_summary", "")).strip():
            continue
        seen.add(case_id)
        verified.append(record)
    return verified


def _verified_document(record: dict[str, Any]) -> str:
    fields = (
        ("사고 유형", record.get("accident_type", "")),
        ("발생일", record.get("accident_date", "")),
        ("업종", record.get("industry", "")),
        ("사고 개요", record.get("accident_summary", "")),
        ("발생 원인", record.get("cause_summary", "")),
        ("예방사항", record.get("prevention_summary", "")),
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
    if VERIFIED_CASE_COLLECTION_NAME in {LAW_COLLECTION_NAME, SOURCE_CASE_COLLECTION_NAME}:
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
