import ast
import json
import re
import unittest
import warnings
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
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", DeprecationWarning)
            import chromadb

        cls.pipeline = PIPELINE_PATH.read_text(encoding="utf-8")
        cls.pipeline_tree = ast.parse(cls.pipeline)
        cls.manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
        cls.ignore = IGNORE_PATH.read_text(encoding="utf-8")
        cls.app = APP_PATH.read_text(encoding="utf-8-sig")
        cls.app_tree = ast.parse(cls.app)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", DeprecationWarning)
            cls.client = chromadb.PersistentClient(path=str(DB_PATH))
            cls.collection = cls.client.get_collection(name="mine_official_accident_cases")
            cls.collection_payload = cls.collection.get(include=["metadatas"])

    def require_integration(self):
        self.assertTrue(DB_PATH.is_dir(), "신규 공식 사고사례 DB가 없습니다.")
        self.assertIn("mine_official_accident_cases", self.app)
        self.assertRegex(self.app, r"def\s+search_official_siren_cases\s*\(")

    def test_01_new_collection_name_exists(self):
        self.assertIn('COLLECTION_NAME = "mine_official_accident_cases"', self.pipeline)
        self.assertEqual("mine_official_accident_cases", self.collection.name)

    def test_02_law_and_case_collections_are_separate(self):
        self.assertIn('LAW_COLLECTION_NAME = "mine_safety_docs"', self.pipeline)
        self.assertIn("COLLECTION_NAME == LAW_COLLECTION_NAME", self.pipeline)
        self.assertIn("23_official_accident_case_vector_db", self.pipeline)
        self.assertIn("10_vector_db_with_major_accident_docs", self.pipeline)
        self.assertNotEqual("mine_official_accident_cases", "mine_safety_docs")

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
        self.assertIn("build_official_case_vector_db", function_names)
        self.assertIn("verify_official_case_vector_db", function_names)
        self.assertIn("filter_cases_for_vector_db", function_names)
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
        self.assertEqual(129, self.collection.count())

    def test_08_missing_db_and_search_failure_fallback_exists(self):
        self.require_integration()
        self.assertRegex(self.app, r"23_official_accident_case_vector_db")
        self.assertRegex(self.app, r"return\s+\[\]")
        self.assertIn("db_unavailable", self.app)
        self.assertIn("search_failed", self.app)

    def test_09_out_of_scope_skips_case_search(self):
        self.require_integration()
        run_flow = next(
            node
            for node in self.app_tree.body
            if isinstance(node, ast.FunctionDef) and node.name == "run_rag_flow"
        )
        out_scope_if = next(
            node
            for node in ast.walk(run_flow)
            if isinstance(node, ast.If)
            and "OUT_OF_SCOPE_INTENT" in ast.unparse(node.test)
            and any(isinstance(item, ast.Return) for item in ast.walk(node))
        )
        search_call = next(
            node
            for node in ast.walk(run_flow)
            if isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "search_official_siren_cases"
        )
        self.assertLess(out_scope_if.lineno, search_call.lineno)

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
        ids = [str(item) for item in self.collection_payload.get("ids", [])]
        self.assertEqual(len(ids), len(set(ids)))

    def test_15_document_and_page_are_rendered(self):
        self.require_integration()
        self.assertIn("source_document", self.app)
        self.assertRegex(self.app, r"page_(?:start|end)")
        metadatas = self.collection_payload.get("metadatas", []) or []
        self.assertEqual(129, len(metadatas))
        self.assertTrue(all(meta.get("source_document") for meta in metadatas))
        self.assertTrue(all(meta.get("page_start") not in (None, "") for meta in metadatas))
        self.assertTrue(all(meta.get("official_case") is True for meta in metadatas))
        self.assertTrue(all(meta.get("mine_relevance") in {"high", "medium"} for meta in metadatas))
        self.assertTrue(all(meta.get("ocr_quality_status") == "pass" for meta in metadatas))
        hashes = [str(meta.get("content_hash", "")) for meta in metadatas]
        self.assertTrue(all(hashes))
        self.assertEqual(len(hashes), len(set(hashes)))

    def test_16_official_case_is_not_a_legal_judgment_notice(self):
        self.require_integration()
        self.assertRegex(self.app, r"법령\s*위반\s*여부.*확정.*근거.*아닙니다")

    def test_17_news_not_official_legal_evidence_notice_remains(self):
        self.assertIn("공식 법령 판단 근거가 아닙니다", self.app)

    def test_18_official_prevention_then_local_fallback_structure(self):
        self.require_integration()
        self.assertIn("prevention_summary", self.app)
        self.assertIn("official_case_warning_points", self.app)
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
