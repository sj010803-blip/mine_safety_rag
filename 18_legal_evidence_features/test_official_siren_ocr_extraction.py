import ast
import importlib.util
import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PIPELINE_PATH = (
    ROOT
    / "21_official_accident_case_pipeline"
    / "build_official_siren_case_db.py"
)
IGNORE_PATH = ROOT / ".gitignore"


class OfficialSirenOcrExtractionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = PIPELINE_PATH.read_text(encoding="utf-8-sig")
        cls.tree = ast.parse(cls.source)
        cls.functions = {
            node.name for node in ast.walk(cls.tree) if isinstance(node, ast.FunctionDef)
        }
        spec = importlib.util.spec_from_file_location("official_siren_ocr_pipeline", PIPELINE_PATH)
        cls.pipeline = importlib.util.module_from_spec(spec)
        assert spec and spec.loader
        spec.loader.exec_module(cls.pipeline)
        cls.ignore = IGNORE_PATH.read_text(encoding="utf-8")

    @staticmethod
    def sample_source():
        return {
            "source_id": "MOEL-SIREN-2026-01-CONSTRUCTION",
            "title": "2026년 1월 중대재해사이렌 건설",
            "coverage_start": "2026-01-01",
            "coverage_end": "2026-01-31",
            "source_page_url": "https://www.moel.go.kr/example",
            "saved_filename": "sample.pdf",
        }

    @staticmethod
    def sample_record(page_number=1, text=None, quality="pass"):
        return {
            "source_id": "MOEL-SIREN-2026-01-CONSTRUCTION",
            "source_file": "sample.pdf",
            "page_number": page_number,
            "quality_status": quality,
            "extracted_text": text or (
                "중대재해 발생 알림\n발생일: 2026년 1월 3일\n"
                "사고 개요: 작업자가 컨베이어 주변을 점검하던 중 끼임 사고가 발생하여 "
                "작업을 중지하고 접근을 통제함\n예방대책: 전원을 차단하고 잠금 조치"
            ),
        }

    def test_01_no_repository_or_backup_recursive_search(self):
        self.assertNotIn("rglob(", self.source)
        self.assertNotIn("os.walk", self.source)
        self.assertNotIn("app_backup_before_", self.source)
        self.assertNotRegex(self.source, r"Get-ChildItem|Select-String|ripgrep")

    def test_02_tesseract_detection_function_exists(self):
        self.assertIn("find_tesseract", self.functions)
        self.assertIn("verify_tesseract_languages", self.functions)

    def test_03_korean_and_english_ocr_language_is_fixed(self):
        self.assertRegex(self.source, r'OCR_LANGUAGE\s*=\s*["\']kor\+eng["\']')
        self.assertIn('{"kor", "eng"}', self.source)

    def test_04_local_tessdata_is_supported(self):
        self.assertIn("select_tessdata_path", self.functions)
        self.assertIn("tessdata_local", self.source)
        self.assertIn("--tessdata-dir", self.source)

    def test_05_page_ocr_cache_structure_exists(self):
        self.assertIn("page_cache_key", self.functions)
        self.assertIn("load_page_ocr_cache", self.functions)
        self.assertIn("cache_page_ocr", self.functions)
        self.assertIn("write_failed_page_records", self.functions)
        self.assertIn("official_siren_failed_pages.jsonl", self.source)
        for field in ("source_sha256", "page_number", "text_hash", "cache_key"):
            self.assertIn(field, self.source)

    def test_06_ocr_retry_is_limited_to_two_attempts(self):
        self.assertEqual(2, self.pipeline.OCR_MAX_ATTEMPTS)
        self.assertIn("attempt_number > OCR_MAX_ATTEMPTS", self.source)

    def test_07_quality_scoring_rejects_empty_and_accepts_korean_text(self):
        empty = self.pipeline.score_ocr_quality("")
        good = self.pipeline.score_ocr_quality(
            "중대재해 발생 알림입니다. 작업자는 즉시 작업을 중지하고 안전한 장소로 대피해야 합니다. "
            "관리자는 위험구역 접근을 통제하고 설비 상태를 다시 점검해야 합니다. " * 2
        )
        self.assertEqual("failed", empty["quality_status"])
        self.assertEqual("pass", good["quality_status"])

    def test_08_empty_ocr_does_not_create_case(self):
        self.assertEqual([], self.pipeline.extract_case_blocks(""))
        valid, reason = self.pipeline.validate_case("")
        self.assertFalse(valid)
        self.assertIn("40자", reason)

    def test_09_cover_and_contents_are_excluded(self):
        self.assertTrue(self.pipeline.is_non_case_page("목차\n1월 사고 목록\n발간사"))
        self.assertEqual([], self.pipeline.extract_case_blocks("목차\n1월 사고 목록\n발간사"))
        divider = "2025 중대재해 사이렌\n사망사고 유발(SIF) 고위험 요인 분석 자료"
        self.assertTrue(self.pipeline.is_non_case_page(divider))

    def test_10_case_id_is_reproducible(self):
        first = self.pipeline.generate_reproducible_case_id(
            "MOEL-SIREN-2026-01-CONSTRUCTION", 1
        )
        second = self.pipeline.generate_reproducible_case_id(
            "MOEL-SIREN-2026-01-CONSTRUCTION", 1
        )
        self.assertEqual(first, second)
        self.assertEqual("KOSHA-SIREN-2026-01-CONSTRUCTION-0001", first)

    def test_11_duplicate_content_hash_is_removed(self):
        source = self.sample_source()
        record = self.sample_record()
        cases, duplicates, _ = self.pipeline.extract_cases([source], [record, dict(record)])
        self.assertEqual(1, len(cases))
        self.assertEqual(1, duplicates)

    def test_12_unclear_ocr_requires_review(self):
        relevance, reason = self.pipeline.classify_mine_relevance(
            "컨베이어 끼임 사고", "review"
        )
        self.assertEqual("review_required", relevance)
        self.assertIn("OCR 품질", reason)

    def test_13_missing_cause_and_prevention_are_not_invented(self):
        source = self.sample_source()
        record = self.sample_record(
            text=(
                "중대재해 발생 알림\n발생일: 2026년 1월 3일\n"
                "사고 개요: 작업 중 사고가 발생하여 작업자가 즉시 대피하고 현장을 통제했습니다. "
                "담당 관리자는 사고 상황과 설비 상태를 확인했습니다."
            )
        )
        case = self.pipeline._parse_case(source, record, record["extracted_text"])
        self.assertEqual("", case["cause_summary"])
        self.assertEqual("", case["prevention_summary"])

    def test_14_pipeline_does_not_import_or_read_app(self):
        self.assertNotRegex(self.source, r"(?:Path|open)\s*\(\s*[\"']app\.py[\"']")
        self.assertNotRegex(self.source, r"APP_(?:PATH|FILE)")
        imported_modules = {
            alias.name
            for node in ast.walk(self.tree)
            if isinstance(node, ast.Import)
            for alias in node.names
        }
        self.assertNotIn("app", imported_modules)

    def test_15_vector_database_is_never_opened(self):
        self.assertNotIn("import chromadb", self.source)
        self.assertNotIn("PersistentClient", self.source)
        with self.assertRaises(self.pipeline.PipelineBlocked):
            self.pipeline.build_vector_db([])

    def test_16_generated_pdf_ocr_and_jsonl_paths_are_ignored(self):
        for expected in (
            "21_official_accident_case_docs/",
            "21_official_accident_case_ocr/",
            "21_official_accident_case_pipeline/tessdata_local/",
            "22_official_accident_case_chunks/",
            "23_official_accident_case_vector_db/",
        ):
            self.assertIn(expected, self.ignore)

    def test_17_no_secret_or_environment_value_access(self):
        self.assertNotRegex(self.source, r"dotenv|load_dotenv|os\.environ|getenv\(")
        self.assertNotRegex(self.source, r"API[_ ]?KEY|CLIENT[_ ]?SECRET")

    def test_18_two_card_layout_is_split_without_third_attempt(self):
        text = (
            "중대재해 발생 알림\n'25년 7월 9일 23:20경 제조 사업장에서 "
            "설비 내부로 떨어져 사망\n예방대책: 추락 방지 조치\n"
            "중대재해 발생 알림\n'25년 9월 15일 09:35경 제조 사업장에서 "
            "부품에 맞은 뒤 난간 사이로 떨어져 사망\n예방대책: 인양계획 준수"
        )
        blocks = self.pipeline.extract_case_blocks(text)
        self.assertEqual(2, len(blocks))
        self.assertTrue(all(self.pipeline.validate_case(block)[0] for block in blocks))
        self.assertEqual(2, self.pipeline.OCR_MAX_ATTEMPTS)
        self.assertEqual(3, self.pipeline.OCR_LAYOUT_PSM)


if __name__ == "__main__":
    unittest.main()
