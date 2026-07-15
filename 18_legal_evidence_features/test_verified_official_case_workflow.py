from __future__ import annotations

import ast
import importlib.util
import json
from pathlib import Path
import re
import tempfile
import unittest
from unittest.mock import patch


ROOT_DIR = Path(__file__).resolve().parents[1]
APP_PATH = ROOT_DIR / "app.py"
PIPELINE_PATH = (
    ROOT_DIR
    / "21_official_accident_case_pipeline"
    / "build_official_siren_case_db.py"
)
REVIEW_MODULE_PATH = ROOT_DIR / "verified_case_review.py"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"모듈을 읽을 수 없습니다: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def function_source(source: str, function_name: str) -> str:
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == function_name:
            return ast.get_source_segment(source, node) or ""
    return ""


class VerifiedOfficialCaseWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app_source = APP_PATH.read_text(encoding="utf-8-sig")
        cls.pipeline_source = PIPELINE_PATH.read_text(encoding="utf-8-sig")
        cls.review_source = REVIEW_MODULE_PATH.read_text(encoding="utf-8-sig")
        cls.review = load_module("verified_case_review_test", REVIEW_MODULE_PATH)
        cls.pipeline = load_module("official_case_pipeline_test", PIPELINE_PATH)

    def test_01_verification_status_structure(self) -> None:
        self.assertEqual(
            set(self.review.VERIFICATION_STATUSES),
            {"unverified", "verified", "rejected", "manual_review"},
        )

    def test_02_existing_case_defaults_to_unverified(self) -> None:
        prepared = self.pipeline.prepare_unverified_review_cases(
            [{"case_id": "CASE-1", "page_start": 3, "accident_summary": "원문"}]
        )
        self.assertEqual(prepared[0]["case_id"], "CASE-1")
        self.assertEqual(prepared[0]["verification_status"], "unverified")
        self.assertEqual(prepared[0]["verified_at"], "")

    def test_03_public_search_requires_verified(self) -> None:
        source = function_source(self.app_source, "search_official_siren_cases")
        self.assertIn('metadata.get("verification_status", "")', source)
        self.assertRegex(source, r'!=\s*["\']verified["\']')

    def test_04_non_verified_statuses_are_excluded(self) -> None:
        records = [
            {"case_id": status, "verification_status": status}
            for status in ("unverified", "manual_review", "rejected")
        ]
        self.assertEqual(self.review.verified_records_only(records), [])

    def test_05_original_page_image_is_required_and_linked(self) -> None:
        source = function_source(self.review_source, "verified_records_only")
        self.assertIn("original_page_image", source)
        self.assertIn("is_file", source)
        self.assertIn("st.image", function_source(self.app_source, "render_verified_official_case_review_page"))

    def test_06_administrator_review_page_exists(self) -> None:
        source = function_source(self.app_source, "render_verified_official_case_review_page")
        self.assertIn("공식 사고사례 검수", source)
        self.assertIn("전체 후보", source)
        self.assertIn("미검증", source)

    def test_07_verified_button_exists(self) -> None:
        source = function_source(self.app_source, "render_verified_official_case_review_page")
        self.assertIn('form_submit_button("검증 완료"', source)
        self.assertIn('target_status = "verified"', source)

    def test_08_manual_edits_can_be_saved(self) -> None:
        source = function_source(self.app_source, "render_verified_official_case_review_page")
        self.assertIn('form_submit_button("수정 저장"', source)
        for field in ("accident_date", "industry", "accident_type", "accident_summary"):
            self.assertIn(f'"{field}"', source)
        self.assertIn("save_review_update", source)

    def test_09_review_history_is_append_only(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            temp_root = Path(temp_name)
            review_dir = temp_root / "review"
            status_path = review_dir / "status.jsonl"
            initial = self.review.default_review_record(
                {
                    "case_id": "CASE-HISTORY",
                    "source_document": "공식 문서",
                    "page_start": 1,
                    "accident_summary": "원문과 대조할 사고 개요",
                }
            )
            with patch.multiple(
                self.review,
                ROOT_DIR=temp_root,
                REVIEW_DIR=review_dir,
                REVIEW_STATUS_PATH=status_path,
            ):
                self.review._append_review_event(initial, "initialized_unverified")
                self.review.save_review_update(
                    "CASE-HISTORY",
                    {"industry": "제조업"},
                    "manual_review",
                    "추가 대조 필요",
                )
                events = [json.loads(line) for line in status_path.read_text(encoding="utf-8").splitlines()]
            self.assertEqual(len(events), 2)
            self.assertEqual(events[-1]["verification_status"], "manual_review")
            self.assertEqual(events[-1]["industry"], "제조업")

    def test_10_original_ocr_jsonl_is_never_overwritten(self) -> None:
        self.assertIn("ORIGINAL_OCR_JSONL_PATH", self.review_source)
        self.assertNotRegex(
            self.review_source,
            r"ORIGINAL_OCR_JSONL_PATH\s*\.\s*(?:write_text|write_bytes|open\s*\([^)]*[wa+])",
        )

    def test_11_verified_collection_is_separate(self) -> None:
        self.assertEqual(
            self.review.VERIFIED_CASE_COLLECTION_NAME,
            "mine_verified_official_accident_cases",
        )

    def test_12_law_collection_remains_separate(self) -> None:
        names = {
            self.review.LAW_COLLECTION_NAME,
            self.review.SOURCE_CASE_COLLECTION_NAME,
            self.review.VERIFIED_CASE_COLLECTION_NAME,
        }
        self.assertEqual(len(names), 3)
        self.assertEqual(self.review.LAW_COLLECTION_NAME, "mine_safety_docs")

    def test_13_unverified_source_collection_remains_separate(self) -> None:
        self.assertEqual(
            self.review.SOURCE_CASE_COLLECTION_NAME,
            "mine_official_accident_cases",
        )
        self.assertNotEqual(
            self.review.SOURCE_CASE_COLLECTION_NAME,
            self.review.VERIFIED_CASE_COLLECTION_NAME,
        )

    def test_14_only_verified_records_feed_verified_db(self) -> None:
        rebuild_source = function_source(self.review_source, "rebuild_verified_case_db")
        self.assertIn("verified_records_only(records)", rebuild_source)
        self.assertIn("verification_status", rebuild_source)
        self.assertIn("verified", rebuild_source)
        self.assertIn('"official_case"', rebuild_source)
        self.assertIn('"ocr_quality_status"', rebuild_source)

    def test_15_source_comparison_badge_is_public(self) -> None:
        source = function_source(self.app_source, "render_official_siren_case_card")
        self.assertIn("원본 PDF와 대조하여 내용 검증이 완료된 공식 사고사례입니다.", source)
        self.assertIn('!= "verified"', source)

    def test_16_out_of_scope_question_still_returns_early(self) -> None:
        source = function_source(self.app_source, "search_official_siren_cases")
        scope_position = source.find("OUT_OF_SCOPE_INTENT")
        collection_position = source.find("load_official_case_collection")
        self.assertGreaterEqual(scope_position, 0)
        self.assertGreater(collection_position, scope_position)
        self.assertIn("return []", source[scope_position:collection_position])

    def test_17_news_and_warning_tabs_remain(self) -> None:
        for label in ("공식 재해사례", "최근 뉴스 참고", "핵심 주의사항"):
            self.assertIn(label, self.app_source)

    def test_18_secrets_are_not_directly_rendered(self) -> None:
        direct_output = re.compile(
            r"st\.(?:write|text|markdown|code|json)\s*\(\s*os\.(?:getenv|environ)"
        )
        self.assertIsNone(direct_output.search(self.app_source))
        self.assertNotRegex(
            self.review_source,
            r'(?:API_KEY|CLIENT_SECRET)\s*=\s*["\'][^"\']+["\']',
        )

    def test_19_broken_ocr_text_is_not_on_public_card(self) -> None:
        public_source = function_source(self.app_source, "render_official_siren_case_card")
        admin_source = function_source(self.app_source, "render_verified_official_case_review_page")
        self.assertNotIn("원문 OCR 내용 보기", public_source)
        self.assertNotIn("layout_ocr_text", public_source)
        self.assertIn("원문 OCR 내용 보기", admin_source)

    def test_20_existing_code_remains_parseable_and_compatible(self) -> None:
        ast.parse(self.app_source)
        ast.parse(self.pipeline_source)
        ast.parse(self.review_source)
        self.assertIn("render_latest_reference_cases", self.app_source)
        self.assertIn("run_rag_flow", self.app_source)


if __name__ == "__main__":
    unittest.main()
