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


class OfficialCaseSearchFallbackAndDbUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app_source = APP_PATH.read_text(encoding="utf-8-sig")
        cls.review_source = REVIEW_PATH.read_text(encoding="utf-8-sig")
        cls.pipeline_source = PIPELINE_PATH.read_text(encoding="utf-8-sig")
        cls.gitignore_source = (ROOT_DIR / ".gitignore").read_text(encoding="utf-8-sig")
        cls.review = load_module("official_case_fallback_review_test", REVIEW_PATH)
        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.temp_root = Path(cls.temp_dir.name)
        (cls.temp_root / "card.png").write_bytes(b"safe image fixture")

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temp_dir.cleanup()

    def text_safe_case(self, case_id: str = "TEXT-1", relation: str = "direct"):
        return {
            "case_id": case_id,
            "content_hash": f"hash-{case_id}",
            "official_case": True,
            "verification_status": "unverified",
            "public_case_tier": "text_safe_fallback",
            "relation_type": relation,
            "source_document": "공식 중대재해사이렌",
            "source_period": "2025",
            "page_start": 3,
            "ocr_quality_status": "pass",
            "display_accident_summary": (
                "작업자가 전기설비를 점검하던 중 충전부에 접촉하여 감전되는 사고가 발생했습니다. "
                "전원 차단 여부를 확인하는 과정에서 발생한 공식 사고사례입니다."
            ),
            "display_cause_summary": "",
            "display_prevention_summary": "",
            "matched_terms": ["감전"],
            "distance": 0.3,
        }

    def auto_case(self, case_id: str, status: str, relation: str = "direct"):
        case = self.text_safe_case(case_id, relation)
        case.update(
            {
                "verification_status": status,
                "public_case_tier": status,
                "original_page_number": 3,
                "original_page_image": "card.png",
                "mine_relevance": "medium",
                "text_quality_score": 95,
                "reading_order_score": 95,
                "metadata_quality_score": 95,
            }
        )
        return case

    def ranked(self, cases):
        with patch.object(self.review, "ROOT_DIR", self.temp_root):
            return self.review.rank_public_official_cases(cases, max_results=3)

    def test_01_search_diagnostic_has_all_pipeline_stages(self):
        source = function_source(self.app_source, "initialize_official_case_search_diagnostic")
        for key in (
            "active_collection_name",
            "collection_total_count",
            "raw_query_result_count",
            "after_duplicate_filter_count",
            "after_text_safety_filter_count",
            "after_verification_tier_filter_count",
            "after_relation_filter_count",
            "final_result_count",
            "question_type",
            "expanded_query_terms",
            "direct_candidate_count",
            "analogous_candidate_count",
            "removal_reasons",
        ):
            self.assertIn(key, source)

    def test_02_collection_count_is_recorded(self):
        source = function_source(self.app_source, "search_official_siren_cases")
        self.assertIn("collection.count()", source)
        self.assertIn('diagnostic["collection_total_count"]', source)

    def test_03_raw_and_final_counts_are_distinct(self):
        source = function_source(self.app_source, "search_official_siren_cases")
        self.assertIn('diagnostic["raw_query_result_count"]', source)
        self.assertIn('"final_result_count": len(selected)', source)

    def test_04_public_case_tiers_are_separate_from_verification_status(self):
        self.assertEqual("text_safe_fallback", self.review.TEXT_SAFE_FALLBACK_TIER)
        self.assertEqual("hidden", self.review.HIDDEN_PUBLIC_TIER)
        self.assertEqual(
            {"verified", "auto_screened", "text_safe_fallback", "hidden"},
            set(self.review.PUBLIC_CASE_TIERS),
        )
        self.assertNotIn("text_safe_fallback", self.review.VERIFICATION_STATUSES)

    def test_05_verified_direct_is_ranked_first(self):
        cases = [
            self.text_safe_case("TEXT"),
            self.auto_case("AUTO", "auto_screened"),
            self.auto_case("VERIFIED", "verified"),
        ]
        self.assertEqual("VERIFIED", self.ranked(cases)[0]["case_id"])

    def test_06_auto_screened_precedes_text_safe(self):
        cases = [self.text_safe_case("TEXT"), self.auto_case("AUTO", "auto_screened")]
        self.assertEqual(["AUTO", "TEXT"], [item["case_id"] for item in self.ranked(cases)])

    def test_07_text_safe_is_public_when_higher_tiers_are_absent(self):
        result = self.ranked([self.text_safe_case()])
        self.assertEqual(["TEXT-1"], [item["case_id"] for item in result])

    def test_08_hidden_manual_and_rejected_cases_are_not_public(self):
        hidden = self.text_safe_case("HIDDEN")
        hidden["public_case_tier"] = "hidden"
        manual = self.text_safe_case("MANUAL")
        manual["verification_status"] = "manual_review"
        rejected = self.text_safe_case("REJECTED")
        rejected["verification_status"] = "rejected"
        self.assertEqual([], self.ranked([hidden, manual, rejected]))

    def test_09_corruption_rules_remain_conservative(self):
        for text in (
            "안전 Zool XSF SCHRHON PuwBct 점검",
            "2025년 1월 1일 사고 2025년 2월 2일 사고",
            "['사고 하나', '사고 둘']",
            "사고 !!!!!!! 예방",
            "반복 반복 반복 반복",
            "작업자가 설비를 점검하던 중 사고가 발생했습니다. 작업자가 설비를 점검하던 중 사고가 발생했습니다.",
        ):
            self.assertTrue(self.review.detect_corrupted_ocr_text(text))

    def test_10_optional_metadata_does_not_block_text_safe_candidate(self):
        case = self.text_safe_case()
        for key in (
            "accident_date",
            "industry",
            "industry_detail",
            "accident_type",
            "display_cause_summary",
            "display_prevention_summary",
            "original_page_image",
        ):
            case.pop(key, None)
        self.assertEqual("", self.review.text_safe_fallback_exclusion_reason(case))

    def test_11_traceability_fields_are_required(self):
        for key in ("case_id", "content_hash", "source_document", "page_start"):
            with self.subTest(key=key):
                case = self.text_safe_case()
                case[key] = ""
                self.assertNotEqual("", self.review.text_safe_fallback_exclusion_reason(case))

    def test_12_direct_analogous_and_broad_family_are_distinct(self):
        direct = self.text_safe_case("DIRECT")
        relation, _ = self.review.classify_public_case_relation(
            direct, ("감전",), ("전원 차단",), "electrical_energy"
        )
        self.assertEqual("direct", relation)
        analogous = self.text_safe_case("ANALOGOUS")
        relation, _ = self.review.classify_public_case_relation(
            analogous, ("컨베이어",), ("감전",), "electrical_energy"
        )
        self.assertEqual("analogous", relation)
        broad = self.text_safe_case("BROAD")
        relation, _ = self.review.classify_public_case_relation(
            broad, ("컨베이어",), ("회전체",), "electrical_energy"
        )
        self.assertEqual("broad_family", relation)

    def test_13_other_risk_family_is_not_returned_by_distance_alone(self):
        electrical = self.text_safe_case()
        relation, _ = self.review.classify_public_case_relation(
            electrical,
            ("컨베이어",),
            ("회전체",),
            "mechanical_entanglement",
        )
        self.assertEqual("", relation)
        generic_lockout = self.text_safe_case("GENERIC-LOCK")
        generic_lockout["display_accident_summary"] = (
            "작업구역의 안전 표지와 출입문 잠금 상태를 확인하던 중 미끄러져 넘어지는 사고가 "
            "발생했습니다. 전기 작업이나 전원 설비와는 관련이 없는 사례입니다."
        ).replace("전기 작업이나 전원 설비와는 관련이 없는 ", "")
        relation, _ = self.review.classify_public_case_relation(
            generic_lockout,
            ("감전",),
            ("누전",),
            "electrical_energy",
        )
        self.assertEqual("", relation)

    def test_14_public_results_are_limited_to_three(self):
        self.assertEqual(3, len(self.ranked([self.text_safe_case(f"T-{i}") for i in range(6)])))

    def test_15_duplicate_case_id_keeps_higher_tier(self):
        text_safe = self.text_safe_case("DUP")
        verified = self.auto_case("DUP", "verified")
        result = self.ranked([text_safe, verified])
        self.assertEqual(1, len(result))
        self.assertEqual("verified", result[0]["verification_status"])

    def test_16_public_card_uses_display_fields(self):
        source = function_source(self.app_source, "render_official_siren_case_card")
        for field in (
            "display_accident_summary",
            "display_cause_summary",
            "display_prevention_summary",
        ):
            self.assertIn(field, source)

    def test_17_public_card_does_not_show_raw_ocr_or_hashes(self):
        source = function_source(self.app_source, "render_official_siren_case_card")
        for marker in (
            "layout_ocr_text",
            "full_accident_summary",
            "text_hash",
            "content_hash",
            "원문 OCR 내용 보기",
        ):
            self.assertNotIn(marker, source)

    def test_18_field_manager_ui_hides_raw_database_folder_name(self):
        self.assertNotIn("VECTOR_DB_DIR.name", self.app_source)
        self.assertNotIn("st.write(db_error)", self.app_source)
        source = function_source(self.app_source, "render_sidebar_database_status")
        self.assertIn("공식 법령 DB: ", source)
        self.assertIn("공식 사례 검색 가능 수", source)

    def test_19_admin_ui_has_human_readable_database_states(self):
        source = function_source(self.app_source, "render_sidebar_database_status")
        for marker in (
            "DB 연결 상태",
            "검증 완료 사례 DB",
            "엄격 자동검사 사례 DB",
            "문자 안전 사례 DB",
        ):
            self.assertIn(marker, source)
        review_page = function_source(self.app_source, "render_verified_official_case_review_page")
        self.assertLess(
            review_page.index("load_official_case_collection.clear()"),
            review_page.index("rebuild_verified_case_db"),
        )
        self.assertIn("except OSError", review_page)

    def test_20_database_paths_are_only_in_admin_expander(self):
        source = function_source(self.app_source, "render_sidebar_database_status")
        self.assertIn('if is_admin:', source)
        self.assertIn('st.sidebar.expander("DB 경로 상세 보기"', source)

    def test_21_long_database_paths_have_safe_wrapping_css(self):
        for rule in ("white-space: normal", "overflow-wrap: anywhere", "word-break: break-all"):
            self.assertIn(rule, self.app_source)
        self.assertIn("db-path-detail", self.app_source)

    def test_22_internal_database_markers_are_not_direct_user_messages(self):
        dangerous = re.compile(
            r"st\.(?:write|text|markdown|error|warning|info)\([^\n]*(?:db_unavailable|collection_missing)"
        )
        self.assertIsNone(dangerous.search(self.app_source))

    def test_23_zero_case_state_and_connection_error_are_distinct(self):
        for marker in (
            "verified_not_created_zero_cases",
            "auto_screened_not_created_zero_cases",
            "text_safe_not_created_zero_cases",
            "verified_db_unavailable",
            "auto_screened_db_unavailable",
            "text_safe_db_unavailable",
        ):
            self.assertIn(marker, self.app_source)

    def test_24_tab_order_is_preserved(self):
        source = function_source(self.app_source, "render_rag_result")
        self.assertIn('["최근 뉴스 참고", "공식 재해사례", "핵심 주의사항"]', source)

    def test_25_news_tab_renderer_is_preserved(self):
        source = function_source(self.app_source, "render_rag_result")
        news_start = source.index("with news_reference_tab:")
        official_start = source.index("with official_case_tab:", news_start)
        self.assertIn("render_live_news_reference_cases", source[news_start:official_start])

    def test_26_warning_points_renderer_and_text_safe_priority_exist(self):
        render_source = function_source(self.app_source, "render_rag_result")
        warning_start = render_source.index("with warning_points_tab:")
        self.assertIn("render_latest_reference_cases", render_source[warning_start:])
        warning_source = function_source(self.app_source, "official_case_warning_points")
        self.assertIn("TEXT_SAFE_FALLBACK_TIER", warning_source)

    def test_27_out_of_scope_returns_before_every_case_database_load(self):
        source = function_source(self.app_source, "search_official_siren_cases")
        guard = source.index("OUT_OF_SCOPE_INTENT")
        for loader in (
            "load_official_case_collection",
            "load_auto_screened_case_collection",
            "load_text_safe_case_collection",
        ):
            self.assertLess(guard, source.index(loader))
        self.assertIn("return []", source[guard:source.index("load_official_case_collection")])

    def test_28_existing_contracts_remain_parseable(self):
        ast.parse(self.app_source)
        ast.parse(self.review_source)
        ast.parse(self.pipeline_source)
        for marker in ("run_rag_flow", "official_case_warning_points", "render_official_siren_cases"):
            self.assertIn(marker, self.app_source)

    def test_29_text_safe_collection_is_separate_from_law_and_case_databases(self):
        names = {
            self.review.LAW_COLLECTION_NAME,
            self.review.SOURCE_CASE_COLLECTION_NAME,
            self.review.VERIFIED_CASE_COLLECTION_NAME,
            self.review.AUTO_SCREENED_CASE_COLLECTION_NAME,
            self.review.TEXT_SAFE_CASE_COLLECTION_NAME,
        }
        self.assertEqual(5, len(names))
        self.assertIn('TEXT_SAFE_COLLECTION_NAME = "mine_text_safe_official_accident_cases"', self.pipeline_source)
        for pattern in (
            "23_text_safe_official_accident_case_vector_db/",
            "23_text_safe_official_accident_case_vector_db_candidate/",
            "23_text_safe_official_accident_case_vector_db_backup_*/",
        ):
            self.assertIn(pattern, self.gitignore_source)
        rebuild_source = function_source(self.review_source, "rebuild_text_safe_case_db")
        self.assertNotIn("PersistentClient(path=str(LAW", rebuild_source)

    def test_30_secret_values_are_not_directly_rendered(self):
        dangerous = re.compile(
            r"st\.(?:write|text|markdown|code|json)\s*\(\s*(?:os\.(?:getenv|environ)|[^\n]*(?:API_KEY|CLIENT_SECRET))"
        )
        self.assertIsNone(dangerous.search(self.app_source))


if __name__ == "__main__":
    unittest.main()
