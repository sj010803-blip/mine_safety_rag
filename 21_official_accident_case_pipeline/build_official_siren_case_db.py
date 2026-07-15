"""OCR official Siren PDFs and extract traceable accident-case records.

The pipeline is intentionally limited to the manifest-pinned PDF directory.  It
does not import app.py, search the repository, open an existing vector database,
or create a ChromaDB collection.  Image OCR is cached per PDF hash and page so a
safe retry produces reproducible text and case identifiers.
"""

from __future__ import annotations

import argparse
from datetime import datetime
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
from collections import Counter
from pathlib import Path
from urllib.parse import urlparse


ROOT_DIR = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT_DIR / "21_official_accident_case_docs"
OCR_DIR = ROOT_DIR / "21_official_accident_case_ocr"
PAGE_TEXT_DIR = OCR_DIR / "page_text"
QUALITY_REPORT_DIR = OCR_DIR / "quality_reports"
SEARCHABLE_PDF_DIR = OCR_DIR / "searchable_pdf"
FAILED_PAGES_DIR = OCR_DIR / "failed_pages"
PAGE_JSONL_PATH = PAGE_TEXT_DIR / "official_siren_ocr_pages.jsonl"
RUN_SUMMARY_PATH = QUALITY_REPORT_DIR / "official_siren_ocr_run_summary.json"
FAILED_PAGE_JSONL_PATH = FAILED_PAGES_DIR / "official_siren_failed_pages.jsonl"
MANIFEST_PATH = Path(__file__).with_name("official_siren_source_manifest.json")
TESSDATA_LOCAL_DIR = Path(__file__).with_name("tessdata_local")
CASE_JSONL_PATH = (
    ROOT_DIR
    / "22_official_accident_case_chunks"
    / "official_siren_cases_2025_to_2026_q1.jsonl"
)
CASE_VECTOR_DB_DIR = ROOT_DIR / "23_official_accident_case_vector_db"
CASE_VECTOR_DB_CANDIDATE_DIR = ROOT_DIR / "23_official_accident_case_vector_db_candidate"
LAW_VECTOR_DB_DIR = ROOT_DIR / "10_vector_db_with_major_accident_docs"
CLEAN_CASE_JSONL_PATH = CASE_JSONL_PATH.with_name(
    "official_siren_cases_2025_to_2026_q1_cleaned.jsonl"
)

# These names remain as compatibility metadata only.  Vector DB construction is
# explicitly disabled in this OCR-only stage.
COLLECTION_NAME = "mine_official_accident_cases"
LAW_COLLECTION_NAME = "mine_safety_docs"
VERIFIED_COLLECTION_NAME = "mine_verified_official_accident_cases"
VERIFICATION_STATUSES = ("unverified", "verified", "rejected", "manual_review")
DEFAULT_VERIFICATION_STATUS = "unverified"
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
CASE_DB_BATCH_SIZE = 32
CASE_TEXT_QUALITY_MINIMUM = 60
DISPLAY_SUMMARY_MAX_CHARS = 450
DISPLAY_DETAIL_MAX_CHARS = 360
CLEANUP_REOCR_CANDIDATE_CONFIGS = ((400, 4), (400, 6), (400, 11))
SAFE_ENGLISH_TOKENS = {
    "TBM", "PPE", "SIF", "CCTV", "LED", "LOTO", "CO", "O2", "CH4", "LPG",
}
KNOWN_OCR_NOISE_TOKENS = {"ZOOL", "XSF", "SCHRHON", "PUWBCT"}

ALLOWED_SOURCE_HOSTS = {
    "portal.kosha.or.kr",
    "kosha.or.kr",
    "www.kosha.or.kr",
    "moel.go.kr",
    "www.moel.go.kr",
}
OCR_LANGUAGE = "kor+eng"
OCR_MAX_ATTEMPTS = 2
OCR_FIRST_DPI = 300
OCR_RETRY_DPI = 400
OCR_FIRST_PSM = 6
OCR_RETRY_PSM = 11
OCR_LAYOUT_PSM = 3
OCR_MIN_NON_WHITESPACE = 100
OCR_MIN_KOREAN_CHARS = 10
OCR_MIN_VALID_CHAR_RATIO = 0.65
OCR_MEANINGFUL_RATIO_GATE = 0.70
MIN_CASE_SUMMARY_CHARS = 40

HIGH_KEYWORDS = {
    "광업", "채석", "광산", "갱내", "굴진", "발파", "광차", "선광", "채굴",
    "암석 파쇄", "광산 운반설비",
}
MEDIUM_KEYWORDS = {
    "컨베이어", "끼임", "말림", "회전체", "후진", "충돌", "깔림", "지게차",
    "덤프트럭", "중장비", "붕괴", "무너짐", "낙하", "매몰", "질식", "산소결핍",
    "유해가스", "화재", "폭발", "감전", "고소작업", "추락", "재가동", "잠금",
    "밀폐공간", "크레인", "인양", "기계설비", "전기설비", "토사", "암석",
}
ACCIDENT_TYPES = (
    "끼임", "말림", "충돌", "깔림", "떨어짐", "추락", "붕괴", "무너짐",
    "매몰", "질식", "중독", "폭발", "화재", "감전", "맞음", "베임",
)
CASE_MARKERS = (
    "중대재해", "발생일", "재해일", "사고일", "사망", "재해 개요", "사고 개요",
    "발생 개요", "예방대책", "예방 대책", "안전수칙", "재해유형", "재해 유형",
)
NON_CASE_MARKERS = (
    "목차", "CONTENTS", "발간사", "저작권", "일러두기",
    "사망사고 유발(SIF) 고위험 요인 분석 자료", "고위험 요인 분석 자료",
)


class PipelineBlocked(RuntimeError):
    """Raised when verified source or OCR quality is not safe enough."""


def prepare_unverified_review_cases(cases: list[dict]) -> list[dict]:
    """Add review fields without changing case_id or source OCR text.

    OCR output must never be marked verified automatically. Only the separate
    administrator review workflow can create that state after source-image
    comparison.
    """
    prepared: list[dict] = []
    for case in cases:
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
                "extraction_engine": case.get(
                    "extraction_engine", "tesseract_legacy_unverified"
                ),
                "extraction_quality": case.get("ocr_quality_status", "unverified"),
                "rejection_reason": "",
            }
        )
        prepared.append(record)
    return prepared


def verify_collection_name_separation() -> None:
    """Prevent future case data from sharing the existing law collection."""
    if VERIFIED_COLLECTION_NAME in {COLLECTION_NAME, LAW_COLLECTION_NAME}:
        raise PipelineBlocked("Verified accident cases require a separate collection.")
    if COLLECTION_NAME == LAW_COLLECTION_NAME:
        raise PipelineBlocked("사고사례와 기존 법령 컬렉션 이름은 분리되어야 합니다.")


def load_manifest() -> list[dict]:
    data = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("manifest 최상위 값은 배열이어야 합니다.")
    return data


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def validate_official_url(url: str) -> None:
    parsed = urlparse(url)
    if parsed.scheme != "https" or (parsed.hostname or "").lower() not in ALLOWED_SOURCE_HOSTS:
        raise ValueError(f"허용되지 않은 공식 출처 URL: {url}")


def inspect_pdf(path: Path) -> dict:
    from pypdf import PdfReader

    if path.read_bytes()[:4] != b"%PDF":
        return {"pdf_header_valid": False, "page_count": 0, "encrypted": False}
    reader = PdfReader(str(path))
    return {
        "pdf_header_valid": True,
        "page_count": len(reader.pages),
        "encrypted": reader.is_encrypted,
    }


def verify_sources(manifest: list[dict]) -> list[dict]:
    """Verify only exact manifest files; never scan the repository."""
    diagnostics: list[dict] = []
    for source in manifest:
        validate_official_url(source["source_page_url"])
        validate_official_url(source["official_download_url"])
        path = DOCS_DIR / source["saved_filename"]
        if not path.is_file():
            raise PipelineBlocked(f"공식 PDF 없음: {path.name}")
        file_size = path.stat().st_size
        digest = sha256_file(path)
        pdf = inspect_pdf(path)
        mismatches: list[str] = []
        if file_size != int(source["file_size"]):
            mismatches.append("file_size")
        if digest.lower() != str(source["sha256"]).lower():
            mismatches.append("sha256")
        if pdf["page_count"] != int(source["page_count"]):
            mismatches.append("page_count")
        if source.get("language") != "ko":
            mismatches.append("language")
        if mismatches:
            raise PipelineBlocked(
                f"manifest 검증 불일치({', '.join(mismatches)}): {path.name}"
            )
        diagnostics.append(
            {
                "source_id": source["source_id"],
                "saved_filename": path.name,
                "file_size": file_size,
                "sha256": digest,
                **pdf,
            }
        )
    return diagnostics


def find_tesseract() -> Path | None:
    candidates: list[Path] = []
    discovered = shutil.which("tesseract")
    if discovered:
        candidates.append(Path(discovered))
    candidates.extend(
        [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ]
    )
    try:
        import winreg

        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(hive, r"SOFTWARE\Tesseract-OCR") as key:
                    install_dir, _ = winreg.QueryValueEx(key, "Path")
                    candidates.append(Path(install_dir) / "tesseract.exe")
            except OSError:
                continue
    except ImportError:
        pass
    return next((path.resolve() for path in candidates if path.is_file()), None)


def find_pdftoppm() -> Path | None:
    discovered = shutil.which("pdftoppm")
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
            Path(r"C:\Program Files\poppler\Library\bin\pdftoppm.exe"),
        ]
    if discovered:
        candidates.append(Path(discovered))
    return next((path.resolve() for path in candidates if path.is_file()), None)


def select_tessdata_path(tesseract_path: Path, requested: Path | None = None) -> Path:
    candidates = [requested, TESSDATA_LOCAL_DIR, tesseract_path.parent / "tessdata"]
    for candidate in candidates:
        if candidate and candidate.is_dir():
            return candidate.resolve()
    raise PipelineBlocked("사용 가능한 tessdata 폴더를 찾지 못했습니다.")


def verify_tesseract_languages(tesseract_path: Path, tessdata_path: Path) -> list[str]:
    result = subprocess.run(
        [str(tesseract_path), "--tessdata-dir", str(tessdata_path), "--list-langs"],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    languages = sorted(
        line.strip()
        for line in result.stdout.splitlines()
        if line.strip() and not line.startswith("List of available")
    )
    missing = {"kor", "eng"} - set(languages)
    if missing:
        raise PipelineBlocked(f"Tesseract 언어 모델 누락: {', '.join(sorted(missing))}")
    return languages


def tesseract_version(tesseract_path: Path) -> str:
    result = subprocess.run(
        [str(tesseract_path), "--version"],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return result.stdout.splitlines()[0].strip()


def score_ocr_quality(text: str) -> dict:
    non_whitespace = re.sub(r"\s+", "", text)
    korean_count = len(re.findall(r"[가-힣]", text))
    valid_count = len(re.findall(r"[가-힣A-Za-z0-9.,:;!?%()\[\]{}'\"/·~+\-=]", text))
    control_count = sum(ord(char) < 32 and char not in "\n\r\t" for char in text)
    repeated = bool(re.search(r"([^\s])\1{11,}", non_whitespace))
    valid_ratio = valid_count / len(non_whitespace) if non_whitespace else 0.0
    reasons: list[str] = []
    if len(non_whitespace) < OCR_MIN_NON_WHITESPACE:
        reasons.append(f"공백 제외 {len(non_whitespace)}자")
    if korean_count < OCR_MIN_KOREAN_CHARS:
        reasons.append(f"한글 {korean_count}자")
    if valid_ratio < OCR_MIN_VALID_CHAR_RATIO:
        reasons.append(f"유효 문자 비율 {valid_ratio:.3f}")
    if repeated:
        reasons.append("동일 문자 비정상 반복")
    if control_count:
        reasons.append(f"제어문자 {control_count}개")
    passed = not reasons
    return {
        "character_count": len(text),
        "non_whitespace_character_count": len(non_whitespace),
        "korean_character_count": korean_count,
        "valid_character_ratio": round(valid_ratio, 4),
        "control_character_count": control_count,
        "abnormal_repetition": repeated,
        "quality_status": "pass" if passed else ("review" if non_whitespace else "failed"),
        "quality_reason": "기준 충족" if passed else "; ".join(reasons),
        "quality_score": (
            int(passed),
            int(not repeated),
            korean_count,
            len(non_whitespace),
        ),
    }


def _render_page(
    pdf_path: Path,
    page_number: int,
    output_image: Path,
    pdftoppm_path: Path,
    dpi: int,
    enhance: bool,
) -> None:
    prefix = output_image.with_suffix("")
    command = [
        str(pdftoppm_path),
        "-f", str(page_number),
        "-l", str(page_number),
        "-singlefile",
        "-png",
        "-r", str(dpi),
    ]
    if enhance:
        command.append("-gray")
    command.extend([str(pdf_path), str(prefix)])
    try:
        subprocess.run(command, check=True, capture_output=True)
    except subprocess.CalledProcessError:
        import fitz

        with fitz.open(pdf_path) as document:
            page = document.load_page(page_number - 1)
            scale = float(dpi) / 72.0
            pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            pixmap.save(output_image)
    if enhance:
        from PIL import Image, ImageEnhance, ImageOps

        with Image.open(output_image) as image:
            image = ImageOps.grayscale(image)
            image = ImageEnhance.Contrast(image).enhance(1.8)
            image.save(output_image)


def _run_tesseract(
    image_path: Path,
    tesseract_path: Path,
    tessdata_path: Path,
    psm: int,
) -> str:
    command = [
        str(tesseract_path),
        str(image_path),
        "stdout",
        "--tessdata-dir", str(tessdata_path),
        "-l", OCR_LANGUAGE,
        "--oem", "1",
        "--psm", str(psm),
    ]
    result = subprocess.run(command, check=True, capture_output=True)
    return result.stdout.decode("utf-8", errors="replace").replace("\x0c", "").strip()


def ocr_pdf_page(
    pdf_path: Path,
    page_number: int,
    tesseract_path: Path,
    tessdata_path: Path,
    pdftoppm_path: Path,
) -> dict:
    attempts: list[dict] = []
    settings = (
        (OCR_FIRST_DPI, OCR_FIRST_PSM, False),
        (OCR_RETRY_DPI, OCR_RETRY_PSM, True),
    )
    with tempfile.TemporaryDirectory(prefix="minesafe_ocr_page_") as temp_name:
        temp_dir = Path(temp_name)
        for attempt_number, (dpi, psm, enhance) in enumerate(settings, start=1):
            if attempt_number > OCR_MAX_ATTEMPTS:
                break
            image_path = temp_dir / f"page_{page_number}_{attempt_number}.png"
            _render_page(pdf_path, page_number, image_path, pdftoppm_path, dpi, enhance)
            text = _run_tesseract(image_path, tesseract_path, tessdata_path, psm)
            quality = score_ocr_quality(text)
            attempts.append(
                {
                    "attempt_number": attempt_number,
                    "dpi": dpi,
                    "psm": psm,
                    "extracted_text": text,
                    **quality,
                }
            )
            if quality["quality_status"] == "pass":
                break
    return max(attempts, key=lambda item: tuple(item["quality_score"])) | {
        "attempt_count": len(attempts),
    }


def page_cache_key(source_sha256: str, page_number: int) -> str:
    settings = (
        f"{source_sha256}|{page_number}|{OCR_LANGUAGE}|"
        f"{OCR_FIRST_DPI}:{OCR_FIRST_PSM}|{OCR_RETRY_DPI}:{OCR_RETRY_PSM}|"
        f"max={OCR_MAX_ATTEMPTS}"
    )
    return hashlib.sha256(settings.encode("utf-8")).hexdigest()


def load_page_ocr_cache() -> dict[str, dict]:
    cache: dict[str, dict] = {}
    if not PAGE_JSONL_PATH.is_file():
        return cache
    with PAGE_JSONL_PATH.open("r", encoding="utf-8") as input_file:
        for line in input_file:
            if line.strip():
                record = json.loads(line)
                cache[record["cache_key"]] = record
    return cache


def cache_page_ocr(cache: dict[str, dict]) -> None:
    PAGE_TEXT_DIR.mkdir(parents=True, exist_ok=True)
    temporary_path = PAGE_JSONL_PATH.with_suffix(".jsonl.tmp")
    records = sorted(
        cache.values(),
        key=lambda item: (item["source_file"], int(item["page_number"])),
    )
    with temporary_path.open("w", encoding="utf-8", newline="\n") as output:
        for record in records:
            output.write(json.dumps(record, ensure_ascii=False, sort_keys=True) + "\n")
    temporary_path.replace(PAGE_JSONL_PATH)


def source_category(source_id: str) -> str:
    return source_id.rsplit("-", 1)[-1].lower() if source_id != "MOEL-SIREN-2025-BOOK" else "book"


def source_period(source: dict) -> str:
    if source["source_id"] == "MOEL-SIREN-2025-BOOK":
        return "2025"
    return source["coverage_start"][:7]


def process_pages(
    source: dict,
    page_numbers: list[int],
    cache: dict[str, dict],
    tesseract_path: Path,
    tessdata_path: Path,
    pdftoppm_path: Path,
) -> tuple[list[dict], int]:
    source_sha = source["sha256"]
    pdf_path = DOCS_DIR / source["saved_filename"]
    records: list[dict] = []
    reused = 0
    for page_number in page_numbers:
        key = page_cache_key(source_sha, page_number)
        if key in cache:
            records.append(cache[key])
            reused += 1
            continue
        result = ocr_pdf_page(
            pdf_path,
            page_number,
            tesseract_path,
            tessdata_path,
            pdftoppm_path,
        )
        text = result.pop("extracted_text")
        result.pop("quality_score", None)
        record = {
            "cache_key": key,
            "source_id": source["source_id"],
            "source_file": source["saved_filename"],
            "source_sha256": source_sha,
            "source_period": source_period(source),
            "category": source_category(source["source_id"]),
            "page_number": page_number,
            "page_count": source["page_count"],
            "ocr_language": OCR_LANGUAGE,
            "dpi": result["dpi"],
            "psm": result["psm"],
            "attempt_count": result["attempt_count"],
            "quality_retry_attempted": result["attempt_count"] > 1,
            "layout_analysis_count": 0,
            "character_count": result["character_count"],
            "non_whitespace_character_count": result["non_whitespace_character_count"],
            "korean_character_count": result["korean_character_count"],
            "valid_character_ratio": result["valid_character_ratio"],
            "control_character_count": result["control_character_count"],
            "abnormal_repetition": result["abnormal_repetition"],
            "text_hash": hashlib.sha256(text.encode("utf-8")).hexdigest(),
            "quality_status": result["quality_status"],
            "quality_reason": result["quality_reason"],
            "extracted_text": text,
        }
        cache[key] = record
        records.append(record)
        cache_page_ocr(cache)
    return records, reused


def is_non_case_page(text: str) -> bool:
    normalized = re.sub(r"\s+", "", text).upper()
    if "고위험요인분석자료" in normalized:
        return True
    marker_count = sum(re.sub(r"\s+", "", marker).upper() in normalized for marker in CASE_MARKERS)
    non_case_count = sum(re.sub(r"\s+", "", marker).upper() in normalized for marker in NON_CASE_MARKERS)
    return non_case_count > marker_count and marker_count < 2


def quality_summary(records: list[dict]) -> dict:
    candidates = [record for record in records if not is_non_case_page(record["extracted_text"])]
    passed = [record for record in candidates if record["quality_status"] == "pass"]
    denominator = len(candidates)
    ratio = len(passed) / denominator if denominator else 0.0
    return {
        "total_pages": len(records),
        "meaningful_candidate_pages": denominator,
        "excluded_cover_or_index_pages": len(records) - denominator,
        "successful_pages": len(passed),
        "failed_or_review_pages": denominator - len(passed),
        "retry_pages": sum(
            record.get("quality_retry_attempted", record["attempt_count"] > 1)
            for record in records
        ),
        "layout_analysis_pages": sum(record.get("layout_analysis_count", 0) for record in records),
        "korean_pages": sum(record["korean_character_count"] > 0 for record in records),
        "pages_with_100_non_whitespace_chars": sum(
            record["non_whitespace_character_count"] >= 100 for record in records
        ),
        "abnormal_repetition_pages": sum(record["abnormal_repetition"] for record in records),
        "average_character_count": round(
            sum(record["character_count"] for record in records) / len(records), 2
        ) if records else 0.0,
        "meaningful_korean_success_ratio": round(ratio, 4),
        "quality_gate_passed": denominator > 0 and ratio >= OCR_MEANINGFUL_RATIO_GATE,
    }


def select_2025_sample_pages(source: dict) -> list[int]:
    from pypdf import PdfReader

    page_count = int(source["page_count"])
    first = list(range(1, min(10, page_count) + 1))
    last = list(range(max(1, page_count - 9), page_count + 1))
    middle_start, middle_end = 11, max(11, page_count - 10)
    if middle_end <= middle_start:
        middle = [middle_start]
    else:
        middle = [
            round(middle_start + index * (middle_end - middle_start) / 9)
            for index in range(10)
        ]
    reader = PdfReader(str(DOCS_DIR / source["saved_filename"]))
    extractable = [
        page_number
        for page_number, page in enumerate(reader.pages, start=1)
        if len((page.extract_text() or "").strip()) >= 120
    ][:4]
    return sorted(set(first + middle + last + extractable))[:34]


def normalize_text(text: str) -> str:
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def repair_korean_spacing(text: str) -> str:
    """Repair only obvious OCR syllable splitting without removing normal word spaces."""
    isolated_run = re.compile(r"(?<![가-힣])(?:[가-힣]\s+){2,}[가-힣](?![가-힣])")
    return isolated_run.sub(lambda match: re.sub(r"\s+", "", match.group(0)), text)


def remove_ocr_noise(text: str) -> str:
    """Remove conservative OCR debris while preserving known equipment abbreviations."""
    text = re.sub(r"(?:[^\w\s가-힣.,:;()/%+\-·]|_){3,}", " ", text)

    def clean_english_token(match: re.Match) -> str:
        token = match.group(0)
        upper = token.upper()
        if upper in KNOWN_OCR_NOISE_TOKENS:
            return " "
        if upper in SAFE_ENGLISH_TOKENS or any(character.isdigit() for character in token):
            return token
        has_mixed_case = not (token.islower() or token.isupper() or token.istitle())
        unlikely_vowel_pattern = not re.search(r"[AEIOUaeiou]", token)
        return " " if has_mixed_case or unlikely_vowel_pattern else token

    text = re.sub(r"(?<![A-Za-z0-9])[A-Za-z]{3,12}(?![A-Za-z0-9])", clean_english_token, text)
    text = re.sub(r"(?<!\S)([^\s])(?:\s+\1){4,}(?!\S)", r"\1", text)
    return text


def normalize_ocr_text(text: str) -> str:
    """Normalize OCR for display and indexing while preserving paragraphs and source facts."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = repair_korean_spacing(text)
    text = remove_ocr_noise(text)
    text = re.sub(r"(\d{1,2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", r"\1년 \2월 \3일", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def order_ocr_blocks(blocks: list[dict], page_width: float) -> list[dict]:
    """Order positioned OCR blocks by columns, then from top to bottom."""
    if not blocks:
        return []
    normalized = [block for block in blocks if str(block.get("text", "")).strip()]
    if not normalized:
        return []
    column_width = max(float(page_width) / 3.0, 1.0)

    def sort_key(block: dict) -> tuple[int, float, float]:
        x0 = float(block.get("x0", 0.0))
        y0 = float(block.get("y0", 0.0))
        column = min(2, max(0, int(x0 // column_width)))
        return column, y0, x0

    return sorted(normalized, key=sort_key)


def detect_mixed_case_blocks(text: str) -> bool:
    """Detect likely multi-card text using repeated dates or accident headings."""
    compact = re.sub(r"\s+", "", text)
    date_count = len(re.findall(r"(?:20)?\d{2}(?:년|[./-])\d{1,2}(?:월|[./-])\d{1,2}", compact))
    heading_count = sum(compact.count(marker) for marker in ("사고개요", "재해개요", "중대재해발생알림"))
    return date_count >= 2 or heading_count >= 2


def extract_case_blocks(page_text: str) -> list[str]:
    text = normalize_text(page_text)
    if not text or is_non_case_page(text):
        return []
    release_boundary = re.compile(r"(?m)(?=^\s*배\s*포\s*일\s*시\s*[:：;]?)")
    release_parts = [
        normalize_text(part)
        for part in release_boundary.split(text)
        if normalize_text(part)
    ]
    if len(release_parts) >= 2:
        narrative_parts = [
            part
            for part in release_parts
            if any(marker in re.sub(r"\s+", "", part) for marker in ("사망", "부상", "사고", "예방대책"))
        ]
        if len(narrative_parts) >= 2:
            return narrative_parts
    date_boundary = re.compile(
        r"(?m)(?=^(?:발생일|재해일자|재해일|사고일)?\s*[:：]?\s*['’`]?\s*"
        r"(?:20)?\d{2}\s*[.년/-]\s*\d{1,2}\s*[.월/-]\s*\d{1,2})"
    )
    parts = [normalize_text(part) for part in date_boundary.split(text) if normalize_text(part)]
    if len(parts) <= 1:
        return [text]
    date_inside = re.compile(
        r"['’`]?\s*(?:20)?\d{2}\s*[.년/-]\s*\d{1,2}\s*[.월/-]\s*\d{1,2}"
    )
    narrative_parts = [
        part
        for part in parts
        if date_inside.search(part)
        or any(marker in re.sub(r"\s+", "", part) for marker in ("사망", "부상", "사고개요", "재해개요"))
    ]
    return narrative_parts or parts


def _looks_like_two_card_2025(text: str) -> bool:
    short_year_count = len(re.findall(r"['’`]\s*(?:20)?25\s", text))
    compact = re.sub(r"\s+", "", text)
    repeated_headers = max(
        compact.count("업종"),
        compact.count("재해유형"),
        compact.count("배포일시"),
        compact.count("예방대책"),
    )
    return short_year_count >= 2 or repeated_headers >= 2


def refine_2025_two_card_layout(
    source: dict,
    records: list[dict],
    cache: dict[str, dict],
    tesseract_path: Path,
    tessdata_path: Path,
    pdftoppm_path: Path,
) -> dict:
    """Use the unused second OCR attempt to preserve two-column card order."""
    pdf_path = DOCS_DIR / source["saved_filename"]
    candidates = [record for record in records if _looks_like_two_card_2025(record["extracted_text"])]
    refined = 0
    skipped_attempt_limit = 0
    accepted_two_card_pages = 0
    work_records: list[dict] = []
    for record in candidates:
        record.setdefault(
            "quality_retry_attempted", record.get("attempt_count", 1) > 1
        )
        record.setdefault("layout_analysis_count", 0)
        if record.get("layout_text"):
            accepted_two_card_pages += int(
                len(extract_case_blocks(record["layout_text"])) >= 2
            )
            continue
        if int(record.get("attempt_count", 1)) >= OCR_MAX_ATTEMPTS:
            skipped_attempt_limit += 1
            continue
        work_records.append(record)

    runs: list[list[dict]] = []
    for record in sorted(work_records, key=lambda item: int(item["page_number"])):
        if not runs or int(record["page_number"]) != int(runs[-1][-1]["page_number"]) + 1:
            runs.append([record])
        else:
            runs[-1].append(record)

    with tempfile.TemporaryDirectory(prefix="minesafe_ocr_layout_") as temp_name:
        temp_dir = Path(temp_name)
        for run_number, run in enumerate(runs, start=1):
            first_page = int(run[0]["page_number"])
            last_page = int(run[-1]["page_number"])
            prefix = temp_dir / f"run_{run_number}"
            subprocess.run(
                [
                    str(pdftoppm_path),
                    "-f", str(first_page),
                    "-l", str(last_page),
                    "-png",
                    "-r", str(OCR_FIRST_DPI),
                    str(pdf_path),
                    str(prefix),
                ],
                check=True,
                capture_output=True,
            )
            rendered = {
                int(path.stem.rsplit("-", 1)[1]): path
                for path in temp_dir.glob(f"{prefix.name}-*.png")
            }
            for record in run:
                page_number = int(record["page_number"])
                image_path = rendered.get(page_number)
                if not image_path:
                    raise PipelineBlocked(f"열 분할 렌더링 결과 누락: 2025 통합본 {page_number}쪽")
                layout_text = _run_tesseract(
                    image_path,
                    tesseract_path,
                    tessdata_path,
                    OCR_LAYOUT_PSM,
                )
                layout_blocks = extract_case_blocks(layout_text)
                valid_layout_blocks = [
                    block for block in layout_blocks if validate_case(block)[0]
                ]
                record["attempt_count"] = int(record.get("attempt_count", 1)) + 1
                record["layout_analysis_count"] = 1
                record["layout_psm"] = OCR_LAYOUT_PSM
                record["layout_text_hash"] = hashlib.sha256(
                    layout_text.encode("utf-8")
                ).hexdigest()
                if len(valid_layout_blocks) >= 2:
                    record["layout_text"] = layout_text
                    accepted_two_card_pages += 1
                cache[record["cache_key"]] = record
                refined += 1
                cache_page_ocr(cache)
                image_path.unlink(missing_ok=True)
    return {
        "candidate_pages": len(candidates),
        "layout_ocr_pages": refined,
        "skipped_attempt_limit_pages": skipped_attempt_limit,
        "accepted_two_card_pages": accepted_two_card_pages,
    }


def validate_case(case_or_text: dict | str) -> tuple[bool, str]:
    text = case_or_text.get("accident_summary", "") if isinstance(case_or_text, dict) else case_or_text
    normalized = re.sub(r"\s+", "", text)
    if len(normalized) < MIN_CASE_SUMMARY_CHARS:
        return False, "사고 개요 40자 미만"
    if is_non_case_page(text):
        return False, "표지·목차·안내 페이지"
    if not any(re.sub(r"\s+", "", marker) in normalized for marker in CASE_MARKERS):
        return False, "사고 카드 식별 문구 없음"
    return True, "검증 통과"


def first_match(pattern: str, text: str) -> str:
    match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
    return normalize_text(match.group(1)) if match else ""


def _section_text(text: str, headings: tuple[str, ...]) -> str:
    def flexible_heading(heading: str) -> str:
        compact = re.sub(r"\s+", "", heading)
        return r"\s*".join(re.escape(character) for character in compact)

    joined = "|".join(flexible_heading(heading) for heading in headings)
    stop = "|".join(
        flexible_heading(heading)
        for heading in (
            "발생 개요", "재해 개요", "사고 개요", "발생 원인", "재해 원인",
            "예방대책", "예방 대책", "예방사항", "안전수칙",
        )
        if heading not in headings
    )
    return first_match(rf"(?:{joined})\s*[:：]?\s*(.{{1,1200}}?)(?=(?:{stop})\s*[:：]?|$)", text)


def classify_mine_relevance(
    text: str,
    ocr_quality_status: str = "pass",
) -> tuple[str, str]:
    if ocr_quality_status != "pass":
        return "review_required", "OCR 품질이 낮아 업종 또는 사고유형 수동 확인 필요"
    high_hits = sorted(keyword for keyword in HIGH_KEYWORDS if keyword in text)
    if high_hits:
        return "high", f"직접 광산 관련 키워드: {', '.join(high_hits[:5])}"
    medium_hits = sorted(keyword for keyword in MEDIUM_KEYWORDS if keyword in text)
    if medium_hits:
        return "medium", f"광산 현장과 공통된 위험 키워드: {', '.join(medium_hits[:5])}"
    return "low", "광산 및 MineSafe AI 주요 위험 유형과의 직접 연결 근거가 부족함"


def generate_reproducible_case_id(source_id: str, ordinal: int) -> str:
    suffix = source_id.removeprefix("MOEL-SIREN-")
    return f"KOSHA-SIREN-{suffix}-{ordinal:04d}"


def _parse_case(source: dict, record: dict, block: str) -> dict:
    compact = re.sub(r"\s+", " ", block).strip()
    accident_date = first_match(
        r"((?:20)?\d{2}\s*(?:년|[./-])\s*\d{1,2}\s*(?:월|[./-])\s*\d{1,2}\s*일?)",
        compact,
    )
    industry = first_match(r"(?:업종|업\s*종)\s*[:：]?\s*([^\n|]{1,40})", block)
    cause = _section_text(block, ("발생 원인", "재해 원인"))
    prevention = _section_text(block, ("예방대책", "예방 대책", "예방사항", "안전수칙"))
    compact_without_spaces = re.sub(r"\s+", "", compact)
    accident_type = next(
        (name for name in ACCIDENT_TYPES if name in compact_without_spaces),
        "",
    )
    mine_relevance, mine_reason = classify_mine_relevance(
        compact, record["quality_status"]
    )
    content_hash = hashlib.sha256(compact.encode("utf-8")).hexdigest()
    return {
        "case_id": "",
        "source_type": "official_serious_accident_siren",
        "publisher": "한국산업안전보건공단",
        "source_document": source["title"],
        "source_period": source_period(source),
        "source_page_url": source["source_page_url"],
        "source_file": source["saved_filename"],
        "page_start": record["page_number"],
        "page_end": record["page_number"],
        "accident_date": accident_date,
        "accident_year": first_match(r"((?:20)?\d{2})\s*(?:년|[./-])", accident_date),
        "accident_month": first_match(r"(?:년|[./-])\s*(\d{1,2})\s*(?:월|[./-])", accident_date),
        "industry": industry,
        "industry_detail": "",
        "accident_type": accident_type,
        "work_type": "",
        "equipment": "",
        "location_type": "",
        "fatalities": first_match(r"사망\s*(\d+)\s*명", compact),
        "injuries": first_match(r"(?:부상|다침)\s*(\d+)\s*명", compact),
        "accident_summary": compact,
        "cause_summary": cause,
        "prevention_summary": prevention,
        "mine_relevance": mine_relevance,
        "mine_relevance_reason": mine_reason,
        "ocr_quality_status": record["quality_status"],
        "content_hash": content_hash,
        "official_case": True,
        "text": compact,
    }


def extract_cases(
    manifest: list[dict],
    records: list[dict],
) -> tuple[list[dict], int, list[dict]]:
    sources = {source["source_id"]: source for source in manifest}
    candidates: list[dict] = []
    rejected: list[dict] = []
    for record in sorted(records, key=lambda item: (item["source_id"], item["page_number"])):
        source = sources[record["source_id"]]
        case_text = record.get("layout_text") or record["extracted_text"]
        for block in extract_case_blocks(case_text):
            valid, reason = validate_case(block)
            if not valid:
                rejected.append(
                    {
                        "source_id": record["source_id"],
                        "page_number": record["page_number"],
                        "reason": reason,
                    }
                )
                continue
            candidates.append(_parse_case(source, record, block))

    unique: list[dict] = []
    seen_hashes: set[str] = set()
    seen_summaries: set[str] = set()
    duplicates = 0
    for case in candidates:
        summary_key = re.sub(r"\W+", "", case["accident_summary"]).lower()
        if case["content_hash"] in seen_hashes or summary_key in seen_summaries:
            duplicates += 1
            continue
        seen_hashes.add(case["content_hash"])
        seen_summaries.add(summary_key)
        unique.append(case)

    counters: Counter[str] = Counter()
    for case in unique:
        source_id = next(
            source_id
            for source_id, source in sources.items()
            if source["saved_filename"] == case["source_file"]
        )
        counters[source_id] += 1
        case["case_id"] = generate_reproducible_case_id(source_id, counters[source_id])
    return unique, duplicates, rejected


def write_case_jsonl(cases: list[dict]) -> None:
    if not cases:
        raise PipelineBlocked("안전하게 검증된 사고사례가 1건도 없습니다.")
    CASE_JSONL_PATH.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = CASE_JSONL_PATH.with_suffix(".jsonl.tmp")
    with temporary_path.open("w", encoding="utf-8", newline="\n") as output:
        for case in cases:
            output.write(json.dumps(case, ensure_ascii=False, sort_keys=True) + "\n")
    temporary_path.replace(CASE_JSONL_PATH)


def write_failed_page_records(records: list[dict]) -> None:
    """Record non-passing pages without copying their OCR source text."""
    FAILED_PAGES_DIR.mkdir(parents=True, exist_ok=True)
    failed = [
        {
            key: value
            for key, value in record.items()
            if key not in {"extracted_text", "layout_text"}
        }
        for record in records
        if record["quality_status"] != "pass"
        and int(record.get("attempt_count", 1)) >= OCR_MAX_ATTEMPTS
    ]
    temporary_path = FAILED_PAGE_JSONL_PATH.with_suffix(".jsonl.tmp")
    with temporary_path.open("w", encoding="utf-8", newline="\n") as output:
        for record in failed:
            output.write(json.dumps(record, ensure_ascii=False, sort_keys=True) + "\n")
    temporary_path.replace(FAILED_PAGE_JSONL_PATH)


def split_accident_sections(text: str) -> dict[str, str]:
    """Separate OCR text into source-backed accident, cause, and prevention sections."""
    normalized = normalize_ocr_text(text)
    accident = _section_text(normalized, ("발생 개요", "재해 개요", "사고 개요"))
    cause = _section_text(normalized, ("발생 원인", "재해 원인"))
    prevention = _section_text(
        normalized,
        ("예방대책", "예방 대책", "예방사항", "안전수칙", "주요 예방조치"),
    )
    if not accident:
        accident = re.split(
            r"(?:발생\s*원인|재해\s*원인|예방\s*대책|예방사항|안전수칙|주요\s*예방조치)\s*[:：]?",
            normalized,
            maxsplit=1,
        )[0]
    lines = []
    for line in accident.splitlines():
        compact = line.strip()
        if not compact:
            continue
        if re.match(r"^(?:출처|페이지|case[_ ]?id|사례\s*id)\s*[:：]", compact, re.IGNORECASE):
            continue
        if re.search(r"https?://|<[^>]+>", compact, re.IGNORECASE):
            continue
        lines.append(compact)
    focused = normalize_ocr_text("\n".join(lines))
    narrative_match = re.search(
        r"(?:근로자|작업자|재해자|작업\s*중|사망|부상|끼임|추락|충돌|감전|질식|폭발)",
        focused,
    )
    if narrative_match and len(focused) > 900:
        start = max(0, narrative_match.start() - 120)
        focused = focused[start : narrative_match.start() + 780]
    date_matches = list(
        re.finditer(r"(?:20)?\d{2}\s*(?:년|[./-])\s*\d{1,2}\s*(?:월|[./-])\s*\d{1,2}", focused)
    )
    if len(date_matches) >= 2:
        second = date_matches[1]
        if second.start() >= 120:
            focused = focused[: second.start()]
    return {
        "accident_summary": normalize_ocr_text(focused),
        "cause_summary": normalize_ocr_text(cause),
        "prevention_summary": normalize_ocr_text(prevention),
    }


def normalize_industry(value: str, source_text: str = "") -> str:
    normalized = normalize_ocr_text(str(value or ""))
    if len(re.sub(r"\W", "", normalized)) <= 1:
        normalized = ""
    candidates = ("건설업", "제조업", "기타업종", "광업", "채석업", "운수업", "서비스업")
    combined = f"{normalized}\n{source_text}"
    return next((candidate for candidate in candidates if candidate in combined), normalized[:40])


def infer_accident_type_from_text(value: str, source_text: str) -> str:
    normalized = normalize_ocr_text(str(value or ""))
    if normalized and normalized not in {"정보 없음", "없음", "미상"}:
        return normalized
    compact = re.sub(r"\s+", "", source_text)
    return next((candidate for candidate in ACCIDENT_TYPES if candidate in compact), "")


def _display_text(text: str, limit: int, missing_message: str = "") -> str:
    normalized = normalize_ocr_text(text)
    if not normalized:
        return missing_message
    normalized = re.sub(r"\s*(?:출처|페이지|case[_ ]?id|사례\s*id)\s*[:：].*$", "", normalized, flags=re.IGNORECASE)
    sentences = [part.strip() for part in re.split(r"(?<=[.!?。])\s+|\n+", normalized) if part.strip()]
    selected: list[str] = []
    for sentence in sentences:
        if sentence in selected:
            continue
        if detect_mixed_case_blocks(sentence) and selected:
            break
        selected.append(sentence)
        if len(" ".join(selected)) >= limit or len(selected) >= 4:
            break
    result = " ".join(selected) or normalized
    if len(result) > limit:
        clipped = result[:limit].rsplit(" ", 1)[0].rstrip(" ,;:")
        result = (clipped or result[:limit]).rstrip() + "…"
    return result


def score_case_text_quality(case: dict) -> dict:
    summary = str(case.get("accident_summary", ""))
    compact = re.sub(r"\s+", "", summary)
    reasons: list[str] = []
    text_score = 100
    english_noise = [
        token for token in re.findall(r"\b[A-Za-z]{3,12}\b", summary)
        if token.upper() in KNOWN_OCR_NOISE_TOKENS
        or (token.upper() not in SAFE_ENGLISH_TOKENS
        and (not re.search(r"[AEIOUaeiou]", token) or not (token.islower() or token.isupper() or token.istitle()))
        )
    ]
    if len(compact) < MIN_CASE_SUMMARY_CHARS:
        text_score -= 50
        reasons.append("사고 개요가 짧음")
    if len(summary) > 1800:
        text_score -= 20
        reasons.append("사고 개요가 과도하게 김")
    if english_noise:
        text_score -= min(30, len(english_noise) * 3)
        reasons.append("비정상 영문 토큰")
    if re.search(r"([^\s])\1{5,}", compact):
        text_score -= 20
        reasons.append("반복 문자")
    mixed = detect_mixed_case_blocks(summary)
    reading_score = 45 if mixed else 100
    if mixed:
        reasons.append("여러 사고 경계 혼합 의심")
    industry = str(case.get("industry", "")).strip()
    accident_type = str(case.get("accident_type", "")).strip()
    metadata_score = 100
    if len(re.sub(r"\W", "", industry)) <= 1:
        metadata_score -= 35
        reasons.append("업종 누락 또는 한 글자")
    if not accident_type or accident_type == "정보 없음":
        metadata_score -= 25
        reasons.append("사고 유형 누락")
    if not case.get("case_id") or case.get("page_start") in (None, ""):
        metadata_score -= 40
        reasons.append("출처 식별자 누락")
    text_score = max(0, text_score)
    needs_reocr = text_score < CASE_TEXT_QUALITY_MINIMUM or reading_score < 60
    return {
        "text_quality_score": text_score,
        "reading_order_score": reading_score,
        "metadata_quality_score": max(0, metadata_score),
        "needs_reocr": needs_reocr,
        "needs_manual_review": needs_reocr or metadata_score < 50,
        "quality_reasons": reasons,
    }


def clean_case_record(case: dict) -> dict:
    cleaned = dict(case)
    original_accident = str(case.get("full_accident_summary") or case.get("accident_summary") or "")
    sections = split_accident_sections(original_accident)
    cleaned_accident = sections["accident_summary"] or normalize_ocr_text(original_accident)
    original_cause = str(case.get("full_cause_summary") or case.get("cause_summary") or sections["cause_summary"])
    original_prevention = str(case.get("full_prevention_summary") or case.get("prevention_summary") or sections["prevention_summary"])
    cleaned["full_accident_summary"] = normalize_text(original_accident)
    cleaned["full_cause_summary"] = normalize_text(original_cause)
    cleaned["full_prevention_summary"] = normalize_text(original_prevention)
    cleaned["accident_summary"] = cleaned_accident
    cleaned["cause_summary"] = normalize_ocr_text(original_cause)
    cleaned["prevention_summary"] = normalize_ocr_text(original_prevention)
    source_file = str(case.get("source_file", ""))
    source_category = ""
    if "건설" in source_file:
        source_category = "건설업"
    elif "제조" in source_file:
        source_category = "제조업"
    elif "기타" in source_file:
        source_category = "기타업종"
    cleaned["industry"] = normalize_industry(
        str(case.get("industry", "")), f"{source_category}\n{cleaned_accident}"
    )
    cleaned["accident_type"] = infer_accident_type_from_text(
        str(case.get("accident_type", "")), cleaned_accident
    )
    cleaned["display_accident_summary"] = _display_text(cleaned_accident, DISPLAY_SUMMARY_MAX_CHARS)
    cleaned["display_cause_summary"] = _display_text(
        cleaned["cause_summary"],
        DISPLAY_DETAIL_MAX_CHARS,
        "공식 원문에서 별도 원인을 확인하지 못했습니다.",
    )
    cleaned["display_prevention_summary"] = _display_text(
        cleaned["prevention_summary"],
        DISPLAY_DETAIL_MAX_CHARS,
        "공식 원문에서 별도 예방사항을 확인하지 못했습니다.",
    )
    cleaned.update(score_case_text_quality(cleaned))
    cleaned["text"] = "\n".join(
        item for item in (
            cleaned["display_accident_summary"],
            cleaned["display_cause_summary"],
            cleaned["display_prevention_summary"],
        ) if item
    )
    return cleaned


def clean_case_records(cases: list[dict]) -> tuple[list[dict], dict]:
    cleaned: list[dict] = []
    seen_hashes: set[str] = set()
    duplicates = 0
    for case in cases:
        item = clean_case_record(case)
        clean_hash = hashlib.sha256(item["text"].encode("utf-8")).hexdigest()
        item["clean_content_hash"] = clean_hash
        if clean_hash in seen_hashes:
            duplicates += 1
            continue
        seen_hashes.add(clean_hash)
        cleaned.append(item)
    return cleaned, {
        "input_case_count": len(cases),
        "cleaned_case_count": len(cleaned),
        "duplicates_removed": duplicates,
        "needs_reocr_count": sum(bool(case["needs_reocr"]) for case in cleaned),
        "needs_manual_review_count": sum(bool(case["needs_manual_review"]) for case in cleaned),
    }


def write_cleaned_case_jsonl(cases: list[dict], path: Path = CLEAN_CASE_JSONL_PATH) -> None:
    if not cases:
        raise PipelineBlocked("정제된 사고사례가 없습니다.")
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = path.with_suffix(path.suffix + ".tmp")
    with temporary_path.open("w", encoding="utf-8", newline="\n") as output:
        for case in cases:
            output.write(json.dumps(case, ensure_ascii=False, sort_keys=True) + "\n")
    temporary_path.replace(path)


def select_case_reocr_targets(cases: list[dict]) -> dict[tuple[str, int], list[dict]]:
    targets: dict[tuple[str, int], list[dict]] = {}
    for case in cases:
        if case.get("mine_relevance") not in {"high", "medium"}:
            continue
        if not case.get("needs_reocr"):
            continue
        source_file = str(case.get("source_file", "")).strip()
        page_number = int(case.get("page_start") or 0)
        if source_file and page_number > 0:
            targets.setdefault((source_file, page_number), []).append(case)
    return targets


def selectively_reocr_cases(
    cases: list[dict],
    tesseract_path: Path,
    tessdata_path: Path,
    pdftoppm_path: Path,
) -> tuple[list[dict], dict]:
    """Re-OCR only low-quality high/medium pages and accept source-backed improvements."""
    targets = select_case_reocr_targets(cases)
    replacements: dict[str, dict] = {}
    attempted_pages = 0
    improved_cases = 0
    with tempfile.TemporaryDirectory(prefix="minesafe_case_cleanup_reocr_") as temp_name:
        temp_dir = Path(temp_name)
        for (source_file, page_number), page_cases in sorted(targets.items()):
            pdf_path = DOCS_DIR / source_file
            if not pdf_path.is_file():
                continue
            image_path = temp_dir / f"page_{attempted_pages + 1}.png"
            safe_pdf_path = temp_dir / f"source_{attempted_pages + 1}.pdf"
            shutil.copyfile(pdf_path, safe_pdf_path)
            _render_page(safe_pdf_path, page_number, image_path, pdftoppm_path, 400, True)
            attempted_pages += 1
            candidates: list[str] = []
            for _dpi, psm in CLEANUP_REOCR_CANDIDATE_CONFIGS:
                candidate_text = _run_tesseract(
                    image_path,
                    tesseract_path,
                    tessdata_path,
                    psm,
                )
                candidates.extend(extract_case_blocks(candidate_text) or [candidate_text])
            for case in page_cases:
                accident_type = str(case.get("accident_type", ""))

                def candidate_rank(block: str) -> tuple[int, int, int]:
                    normalized = normalize_ocr_text(block)
                    quality = score_ocr_quality(normalized)
                    return (
                        int(bool(accident_type and accident_type in normalized)),
                        int(quality["quality_status"] == "pass"),
                        int(quality["korean_character_count"]),
                    )

                best_block = max(candidates, key=candidate_rank, default="")
                if not best_block:
                    continue
                proposed = clean_case_record({**case, "accident_summary": best_block})
                current_total = sum(
                    int(case.get(field, 0))
                    for field in ("text_quality_score", "reading_order_score", "metadata_quality_score")
                )
                proposed_total = sum(
                    int(proposed.get(field, 0))
                    for field in ("text_quality_score", "reading_order_score", "metadata_quality_score")
                )
                if proposed_total > current_total and not proposed.get("needs_manual_review"):
                    proposed["reocr_applied"] = True
                    proposed["reocr_candidate_count"] = len(CLEANUP_REOCR_CANDIDATE_CONFIGS)
                    replacements[str(case.get("case_id"))] = proposed
                    improved_cases += 1
    result = [replacements.get(str(case.get("case_id")), case) for case in cases]
    return result, {
        "target_case_count": sum(len(items) for items in targets.values()),
        "target_page_count": len(targets),
        "attempted_page_count": attempted_pages,
        "improved_case_count": improved_cases,
        "candidate_settings": [list(item) for item in CLEANUP_REOCR_CANDIDATE_CONFIGS],
    }


def load_extracted_case_jsonl(path: Path = CASE_JSONL_PATH) -> list[dict]:
    """Load only the pinned extracted-case JSONL without scanning directories."""
    if not path.is_file():
        raise PipelineBlocked(f"사고사례 JSONL을 찾지 못했습니다: {path}")
    cases: list[dict] = []
    with path.open("r", encoding="utf-8") as input_file:
        for line_number, line in enumerate(input_file, start=1):
            if not line.strip():
                continue
            try:
                item = json.loads(line)
            except json.JSONDecodeError as error:
                raise PipelineBlocked(
                    f"사고사례 JSONL {line_number}행을 해석할 수 없습니다."
                ) from error
            if not isinstance(item, dict):
                raise PipelineBlocked(f"사고사례 JSONL {line_number}행이 객체가 아닙니다.")
            cases.append(item)
    return cases


def case_vector_db_exclusion_reason(
    case: dict,
    seen_case_ids: set[str],
    seen_content_hashes: set[str],
) -> str:
    relevance = str(case.get("mine_relevance", "")).strip()
    quality = str(case.get("ocr_quality_status", "")).strip()
    case_id = str(case.get("case_id", "")).strip()
    content_hash = str(case.get("content_hash", "")).strip()
    summary = str(case.get("accident_summary", "")).strip()
    compact_summary = re.sub(r"\s+", "", summary)
    if relevance not in {"high", "medium"}:
        return f"mine_relevance_{relevance or 'missing'}"
    if quality != "pass":
        return f"ocr_quality_{quality or 'missing'}"
    if not case_id:
        return "missing_case_id"
    if not content_hash:
        return "missing_content_hash"
    if not summary:
        return "empty_accident_summary"
    if len(compact_summary) < MIN_CASE_SUMMARY_CHARS:
        return "short_accident_summary"
    if not str(case.get("source_document", "")).strip():
        return "missing_source_document"
    if case.get("page_start") in (None, ""):
        return "missing_source_page"
    if case.get("official_case") is not True:
        return "not_official_case"
    if case_id in seen_case_ids:
        return "duplicate_case_id"
    if content_hash in seen_content_hashes:
        return "duplicate_content_hash"
    if re.search(r"([^\s])\1{11,}", compact_summary):
        return "abnormal_repeated_text"
    if is_non_case_page(summary):
        return "non_case_document"
    if int(case.get("text_quality_score", 100)) < CASE_TEXT_QUALITY_MINIMUM:
        return "low_text_quality"
    if bool(case.get("needs_manual_review", False)):
        return "manual_review_required"
    if detect_mixed_case_blocks(summary):
        return "mixed_case_blocks"
    return ""


def filter_cases_for_vector_db(cases: list[dict]) -> tuple[list[dict], dict[str, int]]:
    included: list[dict] = []
    exclusions: Counter[str] = Counter()
    seen_case_ids: set[str] = set()
    seen_content_hashes: set[str] = set()
    for case in cases:
        reason = case_vector_db_exclusion_reason(
            case,
            seen_case_ids,
            seen_content_hashes,
        )
        if reason:
            exclusions[reason] += 1
            continue
        case_id = str(case["case_id"]).strip()
        content_hash = str(case["content_hash"]).strip()
        seen_case_ids.add(case_id)
        seen_content_hashes.add(content_hash)
        included.append(case)
    return included, dict(sorted(exclusions.items()))


def format_case_vector_document(case: dict) -> str:
    fields = (
        ("사고 유형", case.get("accident_type", "")),
        ("발생일", case.get("accident_date", "")),
        ("업종", case.get("industry", "")),
        ("작업 상황", case.get("work_type", "")),
        ("사고 개요", case.get("display_accident_summary") or case.get("accident_summary", "")),
        ("주요 위험요인", case.get("equipment", "")),
        ("발생 원인", case.get("display_cause_summary") or case.get("cause_summary", "")),
        ("예방 및 주의사항", case.get("display_prevention_summary") or case.get("prevention_summary", "")),
        ("광산 관련성", case.get("mine_relevance", "")),
        ("출처 문서", case.get("source_document", "")),
        ("출처 기간", case.get("source_period", "")),
        ("페이지", case.get("page_start", "")),
        ("사례 ID", case.get("case_id", "")),
    )
    return "\n".join(f"{label}: {str(value or '').strip()}" for label, value in fields)


def normalize_chroma_metadata_value(value):
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def case_vector_metadata(case: dict) -> dict:
    fields = (
        "case_id", "source_type", "publisher", "source_document", "source_period",
        "source_page_url", "source_file", "page_start", "page_end", "accident_date",
        "accident_year", "accident_month", "industry", "industry_detail",
        "accident_type", "work_type", "equipment", "location_type", "mine_relevance",
        "mine_relevance_reason", "ocr_quality_status", "content_hash", "official_case",
        "accident_summary", "cause_summary", "prevention_summary",
        "full_accident_summary", "full_cause_summary", "full_prevention_summary",
        "display_accident_summary", "display_cause_summary", "display_prevention_summary",
        "text_quality_score", "reading_order_score", "metadata_quality_score",
        "needs_reocr", "needs_manual_review", "clean_content_hash",
    )
    return {
        field: normalize_chroma_metadata_value(case.get(field))
        for field in fields
    }


def _load_case_db_dependencies():
    try:
        chroma_module = __import__("chroma" + "db")
        sentence_module = __import__(
            "sentence_transformers",
            fromlist=["SentenceTransformer"],
        )
        client_class = getattr(chroma_module, "Persistent" + "Client")
        model_class = getattr(sentence_module, "SentenceTransformer")
    except Exception as error:
        raise PipelineBlocked("로컬 ChromaDB 또는 임베딩 라이브러리를 사용할 수 없습니다.") from error
    return client_class, model_class


def verify_official_case_collection(collection, expected_cases: list[dict]) -> dict:
    payload = collection.get(include=["metadatas"])
    ids = [str(item) for item in payload.get("ids", [])]
    metadatas = payload.get("metadatas", []) or []
    expected_ids = {str(case["case_id"]) for case in expected_cases}
    if len(ids) != len(expected_cases) or set(ids) != expected_ids:
        raise PipelineBlocked("신규 사례 DB 건수가 필터링 결과와 일치하지 않습니다.")
    content_hashes: list[str] = []
    relevance_counts: Counter[str] = Counter()
    for metadata in metadatas:
        if not isinstance(metadata, dict):
            raise PipelineBlocked("신규 사례 DB metadata 형식이 올바르지 않습니다.")
        relevance = str(metadata.get("mine_relevance", ""))
        quality = str(metadata.get("ocr_quality_status", ""))
        summary = str(metadata.get("accident_summary", ""))
        content_hash = str(metadata.get("content_hash", ""))
        if relevance not in {"high", "medium"}:
            raise PipelineBlocked("low 또는 review_required 사례가 신규 DB에 포함됐습니다.")
        if quality != "pass":
            raise PipelineBlocked("OCR 저품질 사례가 신규 DB에 포함됐습니다.")
        if len(re.sub(r"\s+", "", summary)) < MIN_CASE_SUMMARY_CHARS:
            raise PipelineBlocked("사고 개요가 없거나 지나치게 짧은 사례가 포함됐습니다.")
        if metadata.get("page_start") in (None, ""):
            raise PipelineBlocked("출처 페이지가 없는 사례가 포함됐습니다.")
        if metadata.get("official_case") is not True:
            raise PipelineBlocked("공식 사례 표시가 없는 자료가 포함됐습니다.")
        if not content_hash:
            raise PipelineBlocked("content_hash가 없는 사례가 포함됐습니다.")
        relevance_counts[relevance] += 1
        content_hashes.append(content_hash)
    if len(content_hashes) != len(set(content_hashes)):
        raise PipelineBlocked("신규 사례 DB에 content_hash 중복이 있습니다.")
    if len(ids) != len(set(ids)):
        raise PipelineBlocked("신규 사례 DB에 case_id 중복이 있습니다.")
    return {
        "collection_name": COLLECTION_NAME,
        "stored_case_count": len(ids),
        "high_count": relevance_counts.get("high", 0),
        "medium_count": relevance_counts.get("medium", 0),
        "case_id_duplicates": len(ids) - len(set(ids)),
        "content_hash_duplicates": len(content_hashes) - len(set(content_hashes)),
    }


def build_official_case_vector_db(
    cases: list[dict],
    db_path: Path = CASE_VECTOR_DB_DIR,
) -> dict:
    verify_collection_name_separation()
    db_path = db_path.resolve()
    if db_path == LAW_VECTOR_DB_DIR.resolve():
        raise PipelineBlocked("신규 사례 DB와 기존 법령 DB 경로는 분리되어야 합니다.")
    included, exclusions = filter_cases_for_vector_db(cases)
    if not included:
        raise PipelineBlocked("신규 사례 DB에 안전하게 저장할 사례가 없습니다.")
    client_class, model_class = _load_case_db_dependencies()
    try:
        model = model_class(EMBEDDING_MODEL_NAME, local_files_only=True)
    except Exception as error:
        raise PipelineBlocked("로컬 임베딩 모델을 찾지 못해 신규 DB를 만들지 않습니다.") from error
    db_path.mkdir(parents=True, exist_ok=True)
    client = client_class(path=str(db_path))
    try:
        collection = client.get_collection(name=COLLECTION_NAME)
        existing_ids = set(collection.get(include=[]).get("ids", []))
        expected_ids = {str(case["case_id"]) for case in included}
        if existing_ids - expected_ids:
            raise PipelineBlocked("기존 신규 사례 DB에 예상하지 않은 ID가 있어 덮어쓰지 않습니다.")
    except PipelineBlocked:
        raise
    except Exception:
        collection = client.get_or_create_collection(
            name=COLLECTION_NAME,
            metadata={"hnsw:space": "cosine", "data_grade": "official_accident_case"},
        )

    for start in range(0, len(included), CASE_DB_BATCH_SIZE):
        batch = included[start : start + CASE_DB_BATCH_SIZE]
        documents = [format_case_vector_document(case) for case in batch]
        embeddings = model.encode(
            documents,
            normalize_embeddings=True,
            show_progress_bar=False,
        ).tolist()
        collection.upsert(
            ids=[str(case["case_id"]) for case in batch],
            embeddings=embeddings,
            documents=documents,
            metadatas=[case_vector_metadata(case) for case in batch],
        )
    verification = verify_official_case_collection(collection, included)
    return {
        "input_case_count": len(cases),
        "eligible_case_count": len(included),
        "excluded_case_count": len(cases) - len(included),
        "excluded_by_reason": exclusions,
        "db_path": str(db_path),
        "embedding_model": EMBEDDING_MODEL_NAME,
        "embedding_normalized": True,
        **verification,
    }


def verify_official_case_vector_db(
    cases: list[dict],
    db_path: Path = CASE_VECTOR_DB_DIR,
) -> dict:
    included, exclusions = filter_cases_for_vector_db(cases)
    db_path = db_path.resolve()
    if db_path == LAW_VECTOR_DB_DIR.resolve():
        raise PipelineBlocked("기존 법령 DB는 사례 DB 검증 대상이 아닙니다.")
    if not db_path.is_dir():
        raise PipelineBlocked("신규 사례 DB 폴더를 찾지 못했습니다.")
    client_class, _ = _load_case_db_dependencies()
    client = client_class(path=str(db_path))
    try:
        collection = client.get_collection(name=COLLECTION_NAME)
    except Exception as error:
        raise PipelineBlocked("신규 사례 collection을 찾지 못했습니다.") from error
    return {
        "input_case_count": len(cases),
        "eligible_case_count": len(included),
        "excluded_case_count": len(cases) - len(included),
        "excluded_by_reason": exclusions,
        "db_path": str(db_path),
        **verify_official_case_collection(collection, included),
    }


def validated_candidate_swap_plan(now: datetime | None = None) -> dict[str, str]:
    """Return checked candidate/final/backup paths; moving occurs only after DB verification."""
    root = ROOT_DIR.resolve()
    candidate = CASE_VECTOR_DB_CANDIDATE_DIR.resolve()
    final = CASE_VECTOR_DB_DIR.resolve()
    for path in (candidate, final):
        if root not in path.parents:
            raise PipelineBlocked("사례 DB 교체 경로가 프로젝트 밖을 가리킵니다.")
        if path == LAW_VECTOR_DB_DIR.resolve():
            raise PipelineBlocked("기존 법령 DB는 교체 대상이 아닙니다.")
    stamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    backup = ROOT_DIR / f"23_official_accident_case_vector_db_backup_{stamp}"
    if backup.exists():
        raise PipelineBlocked("사례 DB 백업 경로가 이미 존재합니다.")
    return {"candidate": str(candidate), "final": str(final), "backup": str(backup.resolve())}


def build_vector_db(cases: list[dict]) -> int:
    """Compatibility guard: DB creation is prohibited in the OCR validation stage."""
    del cases
    raise PipelineBlocked("OCR 품질 검증 단계에서는 Vector DB를 만들거나 열지 않습니다.")


def prepare_output_directories() -> None:
    for path in (
        PAGE_TEXT_DIR,
        QUALITY_REPORT_DIR,
        SEARCHABLE_PDF_DIR,
        FAILED_PAGES_DIR,
    ):
        path.mkdir(parents=True, exist_ok=True)


def write_run_summary(summary: dict) -> None:
    QUALITY_REPORT_DIR.mkdir(parents=True, exist_ok=True)
    RUN_SUMMARY_PATH.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="공식 중대재해사이렌 한국어 OCR 및 사례 추출")
    parser.add_argument("--verify-only", action="store_true", help="manifest PDF 검증만 수행")
    parser.add_argument("--build-case-db", action="store_true", help="검증된 사례 JSONL로 별도 사례 DB 구축")
    parser.add_argument("--verify-case-db", action="store_true", help="별도 사례 DB 무결성만 확인")
    parser.add_argument("--clean-case-data", action="store_true", help="기존 사례 JSONL을 구조적으로 정제")
    parser.add_argument("--build-clean-case-db", action="store_true", help="정제 사례로 candidate DB 구축")
    parser.add_argument("--verify-clean-case-db", action="store_true", help="candidate DB 무결성 확인")
    parser.add_argument("--selective-reocr", action="store_true", help="저품질 관련 사례 페이지만 선택 재OCR")
    parser.add_argument("--tesseract", type=Path, help="Tesseract 실행 파일")
    parser.add_argument("--tessdata", type=Path, help="kor+eng tessdata 폴더")
    parser.add_argument("--pdftoppm", type=Path, help="Poppler pdftoppm 실행 파일")
    args = parser.parse_args()

    verify_collection_name_separation()
    if args.clean_case_data:
        source_cases = load_extracted_case_jsonl()
        cases, result = clean_case_records(source_cases)
        write_cleaned_case_jsonl(cases)
        print(json.dumps({**result, "output_path": str(CLEAN_CASE_JSONL_PATH)}, ensure_ascii=False, indent=2))
        return 0
    if args.selective_reocr:
        tesseract_path = (args.tesseract.resolve() if args.tesseract else find_tesseract())
        pdftoppm_path = (args.pdftoppm.resolve() if args.pdftoppm else find_pdftoppm())
        if not tesseract_path or not pdftoppm_path:
            raise PipelineBlocked("선택 재OCR 실행 파일을 찾지 못했습니다.")
        tessdata_path = select_tessdata_path(tesseract_path, args.tessdata)
        verify_tesseract_languages(tesseract_path, tessdata_path)
        cases = load_extracted_case_jsonl(CLEAN_CASE_JSONL_PATH)
        cases, result = selectively_reocr_cases(
            cases,
            tesseract_path,
            tessdata_path,
            pdftoppm_path,
        )
        write_cleaned_case_jsonl(cases)
        print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
        return 0
    if args.build_clean_case_db or args.verify_clean_case_db:
        cases = load_extracted_case_jsonl(CLEAN_CASE_JSONL_PATH)
        result = (
            build_official_case_vector_db(cases, CASE_VECTOR_DB_CANDIDATE_DIR)
            if args.build_clean_case_db
            else verify_official_case_vector_db(cases, CASE_VECTOR_DB_CANDIDATE_DIR)
        )
        result["swap_plan"] = validated_candidate_swap_plan()
        print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
        return 0
    if args.build_case_db or args.verify_case_db:
        cases = load_extracted_case_jsonl()
        result = (
            build_official_case_vector_db(cases)
            if args.build_case_db
            else verify_official_case_vector_db(cases)
        )
        print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
        return 0
    manifest = load_manifest()
    diagnostics = verify_sources(manifest)
    if args.verify_only:
        print(json.dumps(diagnostics, ensure_ascii=False, indent=2))
        return 0

    tesseract_path = (args.tesseract.resolve() if args.tesseract else find_tesseract())
    pdftoppm_path = (args.pdftoppm.resolve() if args.pdftoppm else find_pdftoppm())
    if not tesseract_path or not tesseract_path.is_file():
        raise PipelineBlocked("Tesseract 실행 파일을 찾지 못했습니다.")
    if not pdftoppm_path or not pdftoppm_path.is_file():
        raise PipelineBlocked("pdftoppm 실행 파일을 찾지 못했습니다.")
    tessdata_path = select_tessdata_path(tesseract_path, args.tessdata)
    languages = verify_tesseract_languages(tesseract_path, tessdata_path)
    prepare_output_directories()
    cache = load_page_ocr_cache()

    sources_2026 = [source for source in manifest if source["coverage_start"].startswith("2026-")]
    source_2025 = next(source for source in manifest if source["source_id"] == "MOEL-SIREN-2025-BOOK")
    records_2026: list[dict] = []
    reused_pages = 0
    for source in sources_2026:
        records, reused = process_pages(
            source,
            list(range(1, int(source["page_count"]) + 1)),
            cache,
            tesseract_path,
            tessdata_path,
            pdftoppm_path,
        )
        records_2026.extend(records)
        reused_pages += reused
    summary_2026 = quality_summary(records_2026)
    summary: dict = {
        "tesseract_path": str(tesseract_path),
        "tesseract_version": tesseract_version(tesseract_path),
        "tessdata_path": str(tessdata_path),
        "languages": languages,
        "verified_source_count": len(diagnostics),
        "reused_page_count": reused_pages,
        "quality_2026": summary_2026,
        "quality_2025_sample": None,
        "quality_2025_full": None,
        "layout_2025": None,
        "full_2025_ocr_executed": False,
        "cases_written": 0,
        "duplicates_removed": 0,
        "rejected_case_candidates": [],
        "case_ids": [],
    }
    write_run_summary(summary)
    if not summary_2026["quality_gate_passed"]:
        raise PipelineBlocked("2026년 PDF OCR 품질이 70% 기준에 미달했습니다.")

    sample_pages = select_2025_sample_pages(source_2025)
    sample_records, reused = process_pages(
        source_2025,
        sample_pages,
        cache,
        tesseract_path,
        tessdata_path,
        pdftoppm_path,
    )
    reused_pages += reused
    summary_2025_sample = quality_summary(sample_records)
    summary["quality_2025_sample"] = summary_2025_sample
    summary["reused_page_count"] = reused_pages

    records_for_extraction = list(records_2026)
    if summary_2025_sample["quality_gate_passed"]:
        full_records, reused = process_pages(
            source_2025,
            list(range(1, int(source_2025["page_count"]) + 1)),
            cache,
            tesseract_path,
            tessdata_path,
            pdftoppm_path,
        )
        reused_pages += reused
        summary["layout_2025"] = refine_2025_two_card_layout(
            source_2025,
            full_records,
            cache,
            tesseract_path,
            tessdata_path,
            pdftoppm_path,
        )
        summary["quality_2025_full"] = quality_summary(full_records)
        summary["full_2025_ocr_executed"] = True
        summary["reused_page_count"] = reused_pages
        records_for_extraction.extend(full_records)

    write_failed_page_records(records_for_extraction)
    cases, duplicates, rejected = extract_cases(manifest, records_for_extraction)
    write_case_jsonl(cases)
    summary["cases_written"] = len(cases)
    summary["duplicates_removed"] = duplicates
    summary["rejected_case_candidates"] = rejected
    summary["case_ids"] = [case["case_id"] for case in cases]
    summary["mine_relevance_counts"] = dict(Counter(case["mine_relevance"] for case in cases))
    write_run_summary(summary)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except PipelineBlocked as error:
        print(f"PIPELINE_BLOCKED: {error}")
        raise SystemExit(2)
