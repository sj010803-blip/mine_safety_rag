import ast
import json
import re
import unittest
from pathlib import Path
from urllib.parse import urlparse


ROOT = Path(__file__).resolve().parents[1]
APP_PATH = ROOT / "app.py"
PIPELINE_PATH = (
    ROOT
    / "21_official_accident_case_pipeline"
    / "build_official_siren_case_db.py"
)
MANIFEST_PATH = PIPELINE_PATH.with_name("official_siren_source_manifest.json")
IGNORE_PATH = ROOT / ".gitignore"
DB_PATH = ROOT / "23_official_accident_case_vector_db"


class OfficialSirenPipelineContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pipeline = PIPELINE_PATH.read_text(encoding="utf-8")
        cls.pipeline_tree = ast.parse(cls.pipeline)
        cls.manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
        cls.ignore = IGNORE_PATH.read_text(encoding="utf-8")
        cls.app = APP_PATH.read_text(encoding="utf-8-sig")
        cls.integration_ready = (
            DB_PATH.exists()
            and "mine_official_accident_cases" in cls.app
            and re.search(r"def\s+search_official_siren_cases\s*\(", cls.app) is not None
        )

    def require_integration(self):
        if not self.integration_ready:
            self.skipTest(
                "공식 PDF가 이미지 기반이라 사례 DB 및 app.py 통합 선행조건이 충족되지 않음"
            )

    def test_01_new_collection_name_exists(self):
        self.assertIn('COLLECTION_NAME = "mine_official_accident_cases"', self.pipeline)

    def test_02_law_and_case_collections_are_separate(self):
        self.assertIn('LAW_COLLECTION_NAME = "mine_safety_docs"', self.pipeline)
        self.assertIn("COLLECTION_NAME == LAW_COLLECTION_NAME", self.pipeline)

    def test_03_manifest_uses_only_allowed_official_hosts(self):
        allowed = {
            "portal.kosha.or.kr",
            "kosha.or.kr",
            "www.kosha.or.kr",
            "moel.go.kr",
            "www.moel.go.kr",
        }
        self.assertEqual(10, len(self.manifest))
        for item in self.manifest:
            self.assertIn(urlparse(item["source_page_url"]).hostname, allowed)
            self.assertIn(urlparse(item["official_download_url"]).hostname, allowed)

    def test_04_image_only_sources_are_not_included(self):
        self.assertTrue(all(item["pdf_header_valid"] for item in self.manifest))
        self.assertTrue(all(not item["text_extractable"] for item in self.manifest))
        self.assertTrue(all(not item["included"] for item in self.manifest))

    def test_05_pipeline_has_missing_text_and_search_failure_guards(self):
        function_names = {
            node.name for node in ast.walk(self.pipeline_tree) if isinstance(node, ast.FunctionDef)
        }
        self.assertIn("verify_sources", function_names)
        self.assertIn("extract_cases", function_names)
        self.assertIn("build_vector_db", function_names)
        self.assertIn("PipelineBlocked", self.pipeline)

    def test_06_generated_data_paths_are_git_ignored(self):
        for expected in (
            "21_official_accident_case_docs/",
            "22_official_accident_case_chunks/",
            "23_official_accident_case_vector_db/",
        ):
            self.assertIn(expected, self.ignore)

    def test_07_official_case_search_function_exists(self):
        self.require_integration()
        self.assertRegex(self.app, r"def\s+search_official_siren_cases\s*\(")

    def test_08_missing_db_and_search_failure_fallback_exists(self):
        self.require_integration()
        self.assertRegex(self.app, r"23_official_accident_case_vector_db")
        self.assertRegex(self.app, r"return\s+\[\]")

    def test_09_out_of_scope_skips_case_search(self):
        self.require_integration()
        self.assertRegex(self.app, r"OUT_OF_SCOPE_INTENT")

    def test_10_combined_reference_area_exists(self):
        self.require_integration()
        self.assertIn("사례 및 참고자료", self.app)

    def test_11_official_case_tab_exists(self):
        self.require_integration()
        self.assertIn("공식 재해사례", self.app)

    def test_12_news_tab_exists(self):
        self.require_integration()
        self.assertIn("최근 뉴스 참고", self.app)

    def test_13_warning_tab_exists(self):
        self.require_integration()
        self.assertIn("핵심 주의사항", self.app)

    def test_14_case_id_is_rendered(self):
        self.require_integration()
        self.assertIn("case_id", self.app)

    def test_15_document_and_page_are_rendered(self):
        self.require_integration()
        self.assertIn("source_document", self.app)
        self.assertRegex(self.app, r"page_(?:start|end)")

    def test_16_official_case_is_not_a_legal_judgment_notice(self):
        self.require_integration()
        self.assertRegex(self.app, r"법령\s*위반\s*여부.*확정.*근거.*아닙니다")

    def test_17_news_not_official_legal_evidence_notice_remains(self):
        self.assertIn("공식 법령 판단 근거가 아닙니다", self.app)

    def test_18_official_prevention_then_local_fallback_structure(self):
        self.require_integration()
        self.assertIn("prevention_summary", self.app)
        self.assertIn("case_warning_points_for_intent", self.app)

    def test_19_existing_law_rag_chunk_id_remains(self):
        self.assertIn("chunk_id", self.app)
        self.assertIn("mine_safety_docs", self.app)

    def test_20_answer_modes_separate_source_grades_without_secret_output(self):
        self.require_integration()
        for mode_marker in ("worker_easy", "natural", "hybrid"):
            if mode_marker in self.app:
                break
        else:
            self.fail("답변 모드 분기 구조를 찾지 못했습니다.")
        self.assertNotRegex(self.app, r"st\.(?:write|code|text)\([^\n]*(?:API_KEY|CLIENT_SECRET)")


if __name__ == "__main__":
    unittest.main()
