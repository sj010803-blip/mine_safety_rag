from __future__ import annotations

import ast
import importlib.util
from pathlib import Path
import re
import tempfile
import unittest
from unittest.mock import patch


ROOT_DIR = Path(__file__).resolve().parents[1]
APP_PATH = ROOT_DIR / "app.py"
REVIEW_PATH = ROOT_DIR / "verified_case_review.py"
PIPELINE_PATH = (
    ROOT_DIR
    / "21_official_accident_case_pipeline"
    / "build_official_siren_case_db.py"
)


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"모듈을 읽을 수 없습니다: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def function_source(source: str, name: str) -> str:
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return ast.get_source_segment(source, node) or ""
    return ""


class AutoScreenedOfficialCaseDisplayTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app_source = APP_PATH.read_text(encoding="utf-8-sig")
        cls.review_source = REVIEW_PATH.read_text(encoding="utf-8-sig")
        cls.pipeline_source = PIPELINE_PATH.read_text(encoding="utf-8-sig")
        cls.review = load_module("auto_screened_review_test", REVIEW_PATH)
        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.temp_root = Path(cls.temp_dir.name)
        (cls.temp_root / "card.png").write_bytes(b"safe source image fixture")

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temp_dir.cleanup()

    def safe_case(self, status: str = "auto_screened", relation: str = "direct", case_id: str = "CASE-1"):
        return {
            "case_id": case_id,
            "content_hash": f"hash-{case_id}",
            "official_case": True,
            "verification_status": status,
            "relation_type": relation,
            "mine_relevance": "medium",
            "ocr_quality_status": "pass",
            "needs_manual_review": False,
            "needs_reocr": False,
            "text_quality_score": 95,
            "reading_order_score": 95,
            "metadata_quality_score": 95,
            "source_document": "공식 중대재해사이렌",
            "page_start": 3,
            "original_page_number": 3,
            "original_page_image": "card.png",
            "industry": "제조업",
            "accident_type": "끼임",
            "display_accident_summary": (
                "근로자가 컨베이어 설비를 점검하던 중 회전체에 끼이는 사고가 발생했습니다. "
                "설비가 완전히 정지되지 않은 상태에서 위험구역에 접근한 사례입니다."
            ),
            "display_cause_summary": "공식 원문에서 별도 원인을 확인하지 못했습니다.",
            "display_prevention_summary": "전원을 차단하고 재가동을 방지한 뒤 점검해야 합니다.",
            "matched_terms": ["컨베이어", "끼임"],
            "distance": 0.2,
        }

    def ranked(self, cases):
        with patch.object(self.review, "ROOT_DIR", self.temp_root):
            return self.review.rank_public_official_cases(cases, max_results=3)

    def test_01_auto_screened_status_exists(self):
        self.assertEqual(self.review.AUTO_SCREENED_STATUS, "auto_screened")
        self.assertIn("auto_screened", self.review.CASE_STATUSES)

    def test_02_verified_is_ranked_before_auto_screened(self):
        auto = self.safe_case("auto_screened", "direct", "AUTO")
        verified = self.safe_case("verified", "direct", "VERIFIED")
        self.assertEqual(["VERIFIED", "AUTO"], [item["case_id"] for item in self.ranked([auto, verified])])

    def test_03_auto_screened_is_used_when_verified_is_absent(self):
        case = self.safe_case("auto_screened", "direct", "AUTO")
        result = self.ranked([case])
        self.assertEqual("AUTO", result[0]["case_id"])
        case["display_cause_summary"] = ""
        case["display_prevention_summary"] = ""
        with patch.object(self.review, "ROOT_DIR", self.temp_root):
            self.assertEqual("", self.review.auto_screened_exclusion_reason(case))

    def test_04_unverified_is_not_public(self):
        self.assertEqual([], self.ranked([self.safe_case("unverified")]))

    def test_05_manual_review_is_not_public(self):
        self.assertEqual([], self.ranked([self.safe_case("manual_review")]))

    def test_06_rejected_is_not_public(self):
        self.assertEqual([], self.ranked([self.safe_case("rejected")]))

    def test_07_direct_and_analogous_relations_are_distinct(self):
        case = self.safe_case()
        direct, terms = self.review.classify_case_relation(case, ("컨베이어",), ("회전체",))
        analogous, analogous_terms = self.review.classify_case_relation(case, ("불발공",), ("회전체",))
        self.assertEqual(("direct", ["컨베이어"]), (direct, terms))
        self.assertEqual(("analogous", ["회전체"]), (analogous, analogous_terms))

    def test_08_direct_auto_precedes_analogous_verified(self):
        direct_auto = self.safe_case("auto_screened", "direct", "DIRECT-AUTO")
        analogous_verified = self.safe_case("verified", "analogous", "ANALOGOUS-VERIFIED")
        result = self.ranked([analogous_verified, direct_auto])
        self.assertEqual(["DIRECT-AUTO", "ANALOGOUS-VERIFIED"], [item["case_id"] for item in result])

    def test_09_corrupted_ocr_detector_blocks_noise_but_keeps_safe_abbreviations(self):
        self.assertTrue(self.review.detect_corrupted_ocr_text("안전 Zool XSF SCHRHON PuwBct 점검"))
        self.assertTrue(self.review.detect_corrupted_ocr_text("2025년 1월 1일 사고 2025년 2월 2일 사고"))
        self.assertTrue(self.review.detect_corrupted_ocr_text("8월 5일 사고 9월 1일 다른 사고"))
        self.assertTrue(self.review.detect_corrupted_ocr_text("['사고 하나', '사고 둘']"))
        self.assertFalse(self.review.detect_corrupted_ocr_text("LOTO와 PPE, CCTV를 확인합니다."))

    def test_10_public_card_reads_display_fields_only(self):
        source = function_source(self.app_source, "render_official_siren_case_card")
        for field in ("display_accident_summary", "display_cause_summary", "display_prevention_summary"):
            self.assertIn(field, source)
        self.assertNotRegex(source, r'case\.get\("(?:accident_summary|cause_summary|prevention_summary|text)"')

    def test_11_raw_ocr_is_not_shown_on_public_card(self):
        public_source = function_source(self.app_source, "render_official_siren_case_card")
        admin_source = function_source(self.app_source, "render_verified_official_case_review_page")
        self.assertNotIn("원문 OCR 내용 보기", public_source)
        self.assertNotIn("layout_ocr_text", public_source)
        self.assertIn("원문 OCR 내용 보기", admin_source)

    def test_12_auto_screened_badge_exists(self):
        source = function_source(self.app_source, "render_official_siren_case_card")
        self.assertIn("자동 품질검사 통과", source)
        self.assertIn("사람의 원문 대조 검수는 아직 완료되지 않았습니다", source)

    def test_13_verified_badge_remains(self):
        source = function_source(self.app_source, "render_official_siren_case_card")
        self.assertIn("원본 PDF와 대조하여 내용 검증이 완료된 공식 사고사례입니다.", source)

    def test_14_analogous_warning_exists(self):
        source = function_source(self.app_source, "render_official_siren_case_card")
        self.assertIn("사고 발생 원리와 위험요인이", source)
        self.assertIn("유사한 공식 사례입니다", source)

    def test_15_auto_screened_collection_is_separate(self):
        self.assertEqual(
            self.review.AUTO_SCREENED_CASE_COLLECTION_NAME,
            "mine_auto_screened_official_accident_cases",
        )

    def test_16_law_collection_is_separate(self):
        self.assertEqual(self.review.LAW_COLLECTION_NAME, "mine_safety_docs")

    def test_17_verified_collection_is_separate(self):
        self.assertEqual(self.review.VERIFIED_CASE_COLLECTION_NAME, "mine_verified_official_accident_cases")

    def test_18_unverified_collection_is_separate(self):
        names = {
            self.review.LAW_COLLECTION_NAME,
            self.review.SOURCE_CASE_COLLECTION_NAME,
            self.review.VERIFIED_CASE_COLLECTION_NAME,
            self.review.AUTO_SCREENED_CASE_COLLECTION_NAME,
        }
        self.assertEqual(4, len(names))
        self.assertEqual(self.review.SOURCE_CASE_COLLECTION_NAME, "mine_official_accident_cases")

    def test_19_search_returns_at_most_three(self):
        cases = [self.safe_case("auto_screened", "direct", f"CASE-{index}") for index in range(5)]
        self.assertEqual(3, len(self.ranked(cases)))
        self.assertIn("OFFICIAL_CASE_TOP_K = 3", self.app_source)

    def test_20_duplicate_case_id_is_removed(self):
        first = self.safe_case("auto_screened", "direct", "DUP")
        second = self.safe_case("verified", "direct", "DUP")
        result = self.ranked([first, second])
        self.assertEqual(1, len(result))
        self.assertEqual("verified", result[0]["verification_status"])

    def test_21_tab_order_is_news_official_warning(self):
        render_source = function_source(self.app_source, "render_rag_result")
        expected = '["최근 뉴스 참고", "공식 재해사례", "핵심 주의사항"]'
        self.assertIn(expected, render_source)

    def test_22_news_tab_keeps_news_renderer(self):
        render_source = function_source(self.app_source, "render_rag_result")
        start = render_source.index("with news_reference_tab:")
        end = render_source.index("with official_case_tab:", start)
        self.assertIn("render_live_news_reference_cases", render_source[start:end])

    def test_23_official_tab_keeps_official_renderer(self):
        render_source = function_source(self.app_source, "render_rag_result")
        start = render_source.index("with official_case_tab:")
        end = render_source.index("with warning_points_tab:", start)
        self.assertIn("render_official_siren_cases", render_source[start:end])

    def test_24_warning_tab_keeps_warning_renderer(self):
        render_source = function_source(self.app_source, "render_rag_result")
        start = render_source.index("with warning_points_tab:")
        self.assertIn("render_latest_reference_cases", render_source[start:])

    def test_25_out_of_scope_returns_before_case_search_and_tabs(self):
        search_source = function_source(self.app_source, "search_official_siren_cases")
        self.assertLess(search_source.index("OUT_OF_SCOPE_INTENT"), search_source.index("load_official_case_collection"))
        render_source = function_source(self.app_source, "render_rag_result")
        scope = render_source.index("OUT_OF_SCOPE_INTENT")
        self.assertLess(scope, render_source.index("st.tabs", scope))
        self.assertIn("return", render_source[scope:render_source.index("st.tabs", scope)])

    def test_26_db_unavailable_is_not_rendered_to_users(self):
        dangerous = re.compile(r"st\.(?:write|text|markdown|error|warning|info)\([^\n]*db_unavailable")
        self.assertIsNone(dangerous.search(self.app_source))
        self.assertIn('"db_unavailable"', self.app_source)

    def test_27_missing_db_and_real_error_have_distinct_states(self):
        for marker in (
            "verified_not_created_zero_cases",
            "auto_screened_not_created_zero_cases",
            "verified_db_unavailable",
            "auto_screened_db_unavailable",
            "no_search_results",
        ):
            self.assertIn(marker, self.app_source)

    def test_28_review_promotion_and_history_remain(self):
        review_page = function_source(self.app_source, "render_verified_official_case_review_page")
        save_source = function_source(self.review_source, "save_review_update")
        self.assertIn('target_status = "verified"', review_page)
        self.assertIn("rebuild_auto_screened_case_db", review_page)
        self.assertIn("_append_review_event", save_source)
        self.assertIn("AUTO_SCREENED_STATUS", self.review_source)

    def test_29_existing_contracts_remain_parseable(self):
        ast.parse(self.app_source)
        ast.parse(self.review_source)
        ast.parse(self.pipeline_source)
        for marker in ("run_rag_flow", "official_case_warning_points", "추가 사례", "mine_official_accident_cases"):
            self.assertIn(marker, self.app_source)

    def test_30_secret_values_are_not_directly_rendered(self):
        dangerous = re.compile(
            r"st\.(?:write|text|markdown|code|json)\s*\(\s*(?:os\.(?:getenv|environ)|[^\n]*(?:API_KEY|CLIENT_SECRET))"
        )
        self.assertIsNone(dangerous.search(self.app_source))


if __name__ == "__main__":
    unittest.main()
