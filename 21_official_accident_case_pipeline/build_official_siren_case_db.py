"""OCR official Siren PDFs and extract traceable accident-case records.

The pipeline is intentionally limited to the manifest-pinned PDF directory.  It
does not import app.py, search the repository, open an existing vector database,
or create a ChromaDB collection.  Image OCR is cached per PDF hash and page so a
safe retry produces reproducible text and case identifiers.
"""

from __future__ import annotations

import argparse
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

# These names remain as compatibility metadata only.  Vector DB construction is
# explicitly disabled in this OCR-only stage.
COLLECTION_NAME = "mine_official_accident_cases"
LAW_COLLECTION_NAME = "mine_safety_docs"

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


def verify_collection_name_separation() -> None:
    """Prevent future case data from sharing the existing law collection."""
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
    candidates = [Path(discovered)] if discovered else []
    candidates.extend(
        [
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
    )
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
    subprocess.run(command, check=True, capture_output=True)
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
    parser.add_argument("--tesseract", type=Path, help="Tesseract 실행 파일")
    parser.add_argument("--tessdata", type=Path, help="kor+eng tessdata 폴더")
    parser.add_argument("--pdftoppm", type=Path, help="Poppler pdftoppm 실행 파일")
    args = parser.parse_args()

    verify_collection_name_separation()
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
