from __future__ import annotations

import ast
import csv
import hashlib
import importlib.util
import json
import re
import struct
import tempfile
import unittest
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[1]
EVAL_DIR = ROOT / "26_rag_retrieval_evaluation"
SCRIPT_PATH = EVAL_DIR / "run_rag_retrieval_quality_evaluation.py"
SOURCE_PATH = ROOT / "02_질문시나리오" / "question_scenarios_110.tsv"
APP_PATH = ROOT / "app.py"
GOLD_PATH = EVAL_DIR / "rag_retrieval_eval_questions_30.tsv"
RESULT_TSV_PATH = EVAL_DIR / "rag_retrieval_eval_results_30.tsv"
RESULT_XLSX_PATH = EVAL_DIR / "rag_retrieval_eval_results_30.xlsx"
MANUAL_XLSX_PATH = EVAL_DIR / "rag_retrieval_manual_review_30.xlsx"
SUMMARY_XLSX_PATH = EVAL_DIR / "rag_retrieval_summary.xlsx"
REPORT_PATH = EVAL_DIR / "rag_retrieval_quality_report.txt"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def _read_tsv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle, delimiter="\t"))


def _load_module():
    spec = importlib.util.spec_from_file_location("minesafe_rag_retrieval_eval", SCRIPT_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("평가 스크립트 모듈 사양을 만들 수 없습니다.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _integrity_values() -> dict[str, str]:
    report = REPORT_PATH.read_text(encoding="utf-8")
    if "[INTEGRITY]" not in report:
        return {}
    values: dict[str, str] = {}
    for line in report.split("[INTEGRITY]", 1)[1].splitlines():
        if "=" in line:
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip()
    return values


def _xlsx_sheet_names(path: Path) -> list[str]:
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    return [node.attrib["name"] for node in root.findall("x:sheets/x:sheet", namespace)]


def _top_level_string_constant(path: Path, name: str) -> str:
    tree = ast.parse(path.read_text(encoding="utf-8-sig"))
    for node in tree.body:
        if isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            value = node.value
            if any(isinstance(target, ast.Name) and target.id == name for target in targets):
                if isinstance(value, ast.Constant) and isinstance(value.value, str):
                    return value.value
    raise AssertionError(f"{name} 상수를 찾지 못했습니다.")


class RagRetrievalQualityEvaluationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.module = _load_module()
        cls.script_source = SCRIPT_PATH.read_text(encoding="utf-8-sig")
        cls.report = REPORT_PATH.read_text(encoding="utf-8")
        cls.integrity = _integrity_values()
        cls.gold_rows = _read_tsv(GOLD_PATH)
        cls.result_rows = _read_tsv(RESULT_TSV_PATH)

    def test_01_original_110_question_file_is_unchanged(self) -> None:
        self.assertEqual(_sha256(SOURCE_PATH), self.module.EXPECTED_SOURCE_SHA256)

    def test_02_gold_set_has_exactly_30_questions(self) -> None:
        self.assertEqual(len(self.gold_rows), 30)

    def test_03_eval_ids_are_r001_through_r030_without_duplicates(self) -> None:
        expected = [f"R{index:03d}" for index in range(1, 31)]
        actual = [row["eval_id"] for row in self.gold_rows]
        self.assertEqual(actual, expected)
        self.assertEqual(len(set(actual)), 30)

    def test_04_source_question_ids_are_unique(self) -> None:
        source_ids = [row["source_question_id"] for row in self.gold_rows]
        self.assertEqual(len(source_ids), len(set(source_ids)))

    def test_05_ten_normalized_categories_have_three_questions_each(self) -> None:
        counts = Counter(row["normalized_category"] for row in self.gold_rows)
        self.assertEqual(len(counts), 10)
        self.assertTrue(all(count == 3 for count in counts.values()))

    def test_06_fixed_seed_selection_is_reproducible(self) -> None:
        source_rows = self.module.read_scenario_rows(SOURCE_PATH)
        first = self.module.select_balanced_questions(source_rows, 20260716)
        second = self.module.select_balanced_questions(source_rows, 20260716)
        self.assertEqual(first, second)
        with self.assertRaises(RuntimeError):
            self.module.select_balanced_questions(source_rows, 1)

    def test_07_frozen_gold_set_refuses_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "gold.tsv"
            self.module.freeze_or_validate_gold_set(path, b"first")
            with self.assertRaises(RuntimeError):
                self.module.freeze_or_validate_gold_set(path, b"second")
            self.assertEqual(path.read_bytes(), b"first")

    def test_08_gold_set_sha256_is_recorded(self) -> None:
        self.assertEqual(_sha256(GOLD_PATH), self.integrity.get("gold_set_sha256"))
        self.assertRegex(_sha256(GOLD_PATH), r"^[0-9A-F]{64}$")

    def test_09_expected_document_patterns_parse_and_match(self) -> None:
        patterns = self.module.build_expected_document_patterns("광산안전법; 광산안전기술기준")
        self.assertEqual(len(patterns), 2)
        match = self.module.match_expected_document("01_광산안전법.txt", patterns)
        self.assertTrue(match["matched"])
        self.assertTrue(match["exact_normalized_match"] or match["title_token_match"])

    def test_10_core_element_terms_parse_with_controlled_aliases(self) -> None:
        terms = self.module.build_expected_core_terms("전원 차단; 작업중지; 잠금표지")
        self.assertEqual(len(terms), 3)
        matched = [self.module.match_core_term("전원 차단 후 작업 중지와 LOTO 실시", term)[0] for term in terms]
        self.assertEqual(matched, [True, True, True])

    def test_11_vector_database_access_is_read_only(self) -> None:
        tree = ast.parse(self.script_source)
        forbidden = {"upsert", "delete", "update", "create_collection", "get_or_create_collection"}
        called_on_collection = {
            node.func.attr
            for node in ast.walk(tree)
            if isinstance(node, ast.Call)
            and isinstance(node.func, ast.Attribute)
            and isinstance(node.func.value, ast.Name)
            and node.func.value.id == "collection"
        }
        self.assertTrue(forbidden.isdisjoint(called_on_collection))
        self.assertIn("query", called_on_collection)
        self.assertIn("client.get_collection", self.script_source)

    def test_12_collection_name_is_mine_safety_docs(self) -> None:
        self.assertEqual(self.module.COLLECTION_NAME, "mine_safety_docs")

    def test_13_collection_count_is_verified_as_820(self) -> None:
        self.assertEqual(self.integrity.get("db_count_before"), "820")
        self.assertEqual(self.integrity.get("db_count_after"), "820")

    def test_14_embedding_model_matches_operational_app_constant(self) -> None:
        app_model = _top_level_string_constant(APP_PATH, "EMBEDDING_MODEL_NAME")
        self.assertEqual(self.module.EMBEDDING_MODEL_NAME, app_model)

    def test_15_embedding_dimension_is_384(self) -> None:
        self.assertEqual(self.module.EMBEDDING_DIMENSION, 384)
        self.assertIn("- 차원: 384", self.report)

    def test_16_retrieval_requests_fixed_top_five(self) -> None:
        class Encoded:
            def tolist(self):
                return [[0.0] * 384]

        class Model:
            def __init__(self):
                self.kwargs = None

            def encode(self, values, **kwargs):
                self.kwargs = kwargs
                return Encoded()

        class Collection:
            def __init__(self):
                self.n_results = None

            def count(self):
                return 10

            def query(self, *, query_embeddings, n_results, include):
                self.n_results = n_results
                return {
                    "documents": [[f"문서 {index}" for index in range(5)]],
                    "metadatas": [[{"source": f"자료 {index}", "chunk_id": f"c{index}"} for index in range(5)]],
                    "distances": [[float(index) for index in range(5)]],
                }

        def make_result(document, metadata, distance, vector_rank):
            return {
                "text": document,
                "source": metadata["source"],
                "chunk_id": metadata["chunk_id"],
                "metadata": metadata,
                "distance": distance,
                "vector_rank": vector_rank,
            }

        namespace = {
            "detect_question_intent": lambda question: "general",
            "expand_search_query": lambda question, intent: question,
            "classify_question_type": lambda question: "general",
            "RERANK_TYPE_CONFIGS": {},
            "make_search_result": make_result,
            "add_type_source_candidates": lambda *args: args[2],
            "rerank_search_results": lambda question, candidates, top_k: candidates[:top_k],
        }
        model = Model()
        collection = Collection()
        hits, _ = self.module.retrieve_operational_top5("질문", collection, model, namespace)
        self.assertEqual(collection.n_results, 5)
        self.assertEqual(len(hits), 5)
        self.assertTrue(model.kwargs["normalize_embeddings"])

    def test_17_document_hit_at_1_calculation(self) -> None:
        metrics = self.module.compute_document_hit_metrics(
            [{"rank": 1, "document_match": {"matched": True}}]
        )
        self.assertEqual(metrics["document_hit_at_1"], 1)

    def test_18_document_hit_at_3_calculation(self) -> None:
        metrics = self.module.compute_document_hit_metrics(
            [{"rank": 3, "document_match": {"matched": True}}]
        )
        self.assertEqual(metrics["document_hit_at_1"], 0)
        self.assertEqual(metrics["document_hit_at_3"], 1)

    def test_19_document_hit_at_5_calculation(self) -> None:
        metrics = self.module.compute_document_hit_metrics(
            [{"rank": 5, "document_match": {"matched": True}}]
        )
        self.assertEqual(metrics["document_hit_at_3"], 0)
        self.assertEqual(metrics["document_hit_at_5"], 1)

    def test_20_reciprocal_rank_calculation(self) -> None:
        metrics = self.module.compute_document_hit_metrics(
            [{"rank": 4, "document_match": {"matched": True}}]
        )
        self.assertAlmostEqual(metrics["reciprocal_rank"], 0.25)

    def test_21_element_coverage_calculation(self) -> None:
        terms = [
            {"term": "작업중지", "aliases": ["작업 중지"]},
            {"term": "대피", "aliases": ["대피"]},
        ]
        hits = [{"document_text": "즉시 작업 중지"}, {"document_text": "안전한 곳으로 대피"}]
        coverage_1, _, _ = self.module.compute_element_coverage(hits, terms, 1)
        coverage_2, matched, _ = self.module.compute_element_coverage(hits, terms, 2)
        self.assertEqual(coverage_1, 0.5)
        self.assertEqual(coverage_2, 1.0)
        self.assertEqual(len(matched), 2)

    def test_22_retrieval_diversity_calculation(self) -> None:
        hits = [{"document_title": title} for title in ["A", "A", "B", "B", "C"]]
        diversity = self.module.compute_retrieval_diversity(hits)
        self.assertEqual(diversity["unique_document_count_top5"], 3)
        self.assertEqual(diversity["duplicate_document_count_top5"], 2)
        self.assertAlmostEqual(diversity["same_document_ratio"], 0.4)

    def test_23_metadata_completeness_calculation(self) -> None:
        hits = [
            {"document_title": "A", "chunk_id": "c1", "source": "A", "page": "1"},
            {"document_title": "B", "chunk_id": "c2", "source": "B", "page": ""},
        ]
        self.assertAlmostEqual(self.module.compute_metadata_completeness(hits), 7 / 8)

    def test_24_distance_is_preserved_as_diagnostic_not_percentage(self) -> None:
        hit = {
            "rank": 1,
            "source": "자료",
            "chunk_id": "c1",
            "text": "작업 중지",
            "distance": 1.2345,
            "metadata": {"source": "자료", "chunk_id": "c1"},
        }
        enriched = self.module.enrich_retrieval_hit(hit, [], [])
        self.assertEqual(enriched["distance"], 1.2345)
        self.assertNotRegex(self.script_source, r"distance\s*\*\s*100")

    def test_25_automatic_review_rule_flags_weak_retrieval(self) -> None:
        status, reason = self.module.decide_auto_review(
            [{"raw": "기대 문서"}],
            {
                "document_hit_at_3": 0,
                "document_hit_at_5": 0,
                "element_coverage_at_3": 0.1,
                "element_coverage_at_5": 0.2,
                "metadata_completeness_top5": 0.75,
            },
            {"unique_document_count_top5": 3},
        )
        self.assertEqual(status, "REVIEW")
        self.assertIn("Hit@5=0", reason)

    def test_26_manual_review_cells_are_blank_with_validations(self) -> None:
        namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        with ZipFile(MANUAL_XLSX_PATH) as archive:
            root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        validations = root.findall("x:dataValidations/x:dataValidation", namespace)
        self.assertEqual(len(validations), 2)
        self.assertEqual({item.attrib["sqref"] for item in validations}, {"N2:N151", "O2:O151"})
        for cell in root.findall(".//x:c", namespace):
            match = re.match(r"^[NOP](\d+)$", cell.attrib.get("r", ""))
            if match and int(match.group(1)) >= 2:
                value = cell.find("x:v", namespace)
                inline = cell.find("x:is/x:t", namespace)
                self.assertFalse((value is not None and value.text) or (inline is not None and inline.text))

    def test_27_result_tsv_has_required_columns_and_30_rows(self) -> None:
        self.assertEqual(len(self.result_rows), 30)
        self.assertTrue(set(self.module.RESULT_COLUMNS).issubset(self.result_rows[0]))

    def test_28_result_workbook_has_ten_required_sheets_and_valid_xml(self) -> None:
        expected = [
            "00_요약", "01_30문항_질문세트", "02_질문별_검색결과", "03_Top5_상세",
            "04_유형별_결과", "05_난이도별_결과", "06_문서별_검색빈도",
            "07_핵심요소_포함률", "08_수동검토_대상", "09_평가방법_주의사항",
        ]
        self.assertEqual(_xlsx_sheet_names(RESULT_XLSX_PATH), expected)
        formula_tag = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f"
        namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        with ZipFile(RESULT_XLSX_PATH) as archive:
            category_sheet = ET.fromstring(archive.read("xl/worksheets/sheet5.xml"))
        category_header = category_sheet.find(".//x:c[@r='A1']/x:v", namespace)
        self.assertIsNotNone(category_header)
        self.assertEqual(category_header.text, "normalized_category")
        for workbook in (RESULT_XLSX_PATH, MANUAL_XLSX_PATH, SUMMARY_XLSX_PATH):
            with ZipFile(workbook) as archive:
                for name in archive.namelist():
                    if name.endswith(".xml"):
                        root = ET.fromstring(archive.read(name))
                        self.assertFalse(root.findall(f".//{formula_tag}"))

    def test_29_three_png_charts_exist_with_useful_resolution(self) -> None:
        names = [
            "rag_retrieval_hit_at_k.png",
            "rag_retrieval_category_hit3.png",
            "rag_retrieval_difficulty_hit3.png",
        ]
        for name in names:
            data = (EVAL_DIR / name).read_bytes()
            self.assertEqual(data[:8], b"\x89PNG\r\n\x1a\n")
            width, height = struct.unpack(">II", data[16:24])
            self.assertGreaterEqual(width, 800)
            self.assertGreaterEqual(height, 500)

    def test_30_app_py_is_unchanged(self) -> None:
        self.assertEqual(_sha256(APP_PATH), self.module.EXPECTED_APP_SHA256)

    def test_31_vector_database_logical_snapshot_is_unchanged(self) -> None:
        self.assertEqual(self.integrity.get("db_count_before"), self.integrity.get("db_count_after"))
        self.assertEqual(
            self.integrity.get("db_ids_sha256_before"),
            self.integrity.get("db_ids_sha256_after"),
        )

    def test_32_existing_chunks_are_not_accessed_or_modified(self) -> None:
        self.assertNotIn("08_chunks", self.script_source)
        self.assertNotIn("chunks_with_major_accident_docs", self.script_source)

    def test_33_report_records_unchanged_question_source_sha256(self) -> None:
        current = _sha256(SOURCE_PATH)
        self.assertEqual(self.integrity.get("source_questions_before_sha256"), current)
        self.assertEqual(self.integrity.get("source_questions_after_sha256"), current)

    def test_34_report_distinguishes_three_evaluation_concepts(self) -> None:
        numbered_sections = [
            int(match.group(1))
            for match in re.finditer(r"(?m)^(\d+)\.\s", self.report)
        ]
        self.assertEqual(numbered_sections, list(range(1, 44)))
        self.assertIn("기능 회귀 테스트", self.report)
        self.assertIn("RAG 검색 적합성 평가", self.report)
        self.assertIn("답변 정확성 평가", self.report)
        self.assertIn("법적 정확도", self.report)

    def test_35_no_secret_or_environment_value_access(self) -> None:
        lowered = self.script_source.lower()
        for forbidden in ("load_dotenv", "os.getenv", "os.environ", "api_key", "client_secret"):
            self.assertNotIn(forbidden, lowered)


if __name__ == "__main__":
    unittest.main()
