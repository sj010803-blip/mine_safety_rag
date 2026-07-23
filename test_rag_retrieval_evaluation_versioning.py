from __future__ import annotations

import csv
import hashlib
import json
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parent
BASELINE_DIR = ROOT / "26_rag_retrieval_evaluation"
POSTFIX_DIR = ROOT / "27_rag_retrieval_evaluation_postfix"
MANIFEST_PATH = POSTFIX_DIR / "evaluation_manifest.json"
GENERAL_QUERY_DIR = ROOT / "28_general_query_understanding_evaluation"
GENERAL_QUERY_MANIFEST_PATH = GENERAL_QUERY_DIR / "evaluation_manifest.json"
APP_PATH = ROOT / "app.py"
SOURCE_PATH = ROOT / "02_질문시나리오" / "question_scenarios_110.tsv"
GOLD_PATH = BASELINE_DIR / "rag_retrieval_eval_questions_30.tsv"
RESULT_PATH = POSTFIX_DIR / "rag_retrieval_postfix_results_30.tsv"

EXPECTED_BASELINE_APP_SHA256 = "983B80A2C390829BA7750830696BC333E8AEABC1A7A4F7407237995AB4AC2FF2"
EXPECTED_POSTFIX_APP_SHA256 = "3D798DA421F8BBE08B16135C85E084ACB85D0ACD16AC8939A647A184A27E4680"
EXPECTED_GENERAL_QUERY_APP_SHA256 = "EA648E523D419145B47575D7B7334FCE4963E857914B7FF411E87190452107DA"
EXPECTED_SOURCE_SHA256 = "544888A7717DD31CB0AF5099128D8493DE7AD50EB38F9ACE90FEFA2AFD72277A"
EXPECTED_GOLD_SHA256 = "F950A1505743A4B3607F0EC3BC60D8ED477FB6C3627F14FCD156BF7C81B7D076"

REQUIRED_BASELINE_FILES = {
    "rag_retrieval_eval_questions_30.tsv",
    "rag_retrieval_eval_results_30.tsv",
    "rag_retrieval_eval_results_30.xlsx",
    "rag_retrieval_manual_review_30.xlsx",
    "rag_retrieval_summary.xlsx",
    "rag_retrieval_quality_report.txt",
    "run_rag_retrieval_quality_evaluation.py",
}
REQUIRED_POSTFIX_FILES = {
    "run_postfix_rag_retrieval_evaluation.py",
    "build_postfix_workbooks.mjs",
    "rag_retrieval_postfix_results_30.tsv",
    "rag_retrieval_postfix_results_30.xlsx",
    "rag_retrieval_postfix_manual_review_30.xlsx",
    "rag_retrieval_postfix_summary.xlsx",
    "rag_retrieval_baseline_vs_postfix.xlsx",
    "rag_retrieval_baseline_vs_postfix.png",
    "rag_retrieval_postfix_quality_report.txt",
}
REQUIRED_GENERAL_QUERY_FILES = {
    "development_question_matrix.jsonl",
    "holdout_question_matrix.jsonl",
    "development_question_results.tsv",
    "holdout_question_results.tsv",
    "general_query_understanding_summary.json",
    "rag_retrieval_regression_results_30.tsv",
    "general_query_understanding_quality_report.txt",
    "run_general_query_understanding_evaluation.py",
}
REQUIRED_SCORE_FIELDS = [
    "document_hit_at_1",
    "document_hit_at_3",
    "document_hit_at_5",
    "reciprocal_rank",
    "element_coverage_at_1",
    "element_coverage_at_3",
    "element_coverage_at_5",
    "metadata_completeness_top5",
]


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def read_tsv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle, delimiter="\t"))


class RagRetrievalEvaluationVersioningTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
        cls.general_query_manifest = json.loads(
            GENERAL_QUERY_MANIFEST_PATH.read_text(encoding="utf-8")
        )
        cls.gold_rows = read_tsv(GOLD_PATH)
        cls.result_rows = read_tsv(RESULT_PATH)

    def test_01_baseline_major_files_exist(self) -> None:
        actual = {path.name for path in BASELINE_DIR.iterdir() if path.is_file()}
        self.assertTrue(REQUIRED_BASELINE_FILES.issubset(actual))

    def test_02_baseline_file_hashes_match_manifest(self) -> None:
        recorded = self.manifest["baseline_files_sha256"]
        self.assertTrue(REQUIRED_BASELINE_FILES.issubset(recorded))
        for name, expected in recorded.items():
            self.assertEqual(sha256(BASELINE_DIR / name), expected, name)

    def test_03_baseline_app_sha_is_recorded(self) -> None:
        self.assertEqual(
            self.manifest["baseline_app_sha256"],
            EXPECTED_BASELINE_APP_SHA256,
        )

    def test_04_postfix_manifest_remains_historical_and_current_app_matches_28(self) -> None:
        self.assertEqual(self.manifest["postfix_app_sha256"], EXPECTED_POSTFIX_APP_SHA256)
        self.assertEqual(
            self.general_query_manifest["previous_app_sha256"],
            self.manifest["postfix_app_sha256"],
        )
        self.assertEqual(
            self.general_query_manifest["current_app_sha256"],
            EXPECTED_GENERAL_QUERY_APP_SHA256,
        )
        self.assertEqual(sha256(APP_PATH), self.general_query_manifest["current_app_sha256"])

    def test_05_original_question_set_sha_is_unchanged(self) -> None:
        self.assertEqual(self.manifest["original_question_set_sha256"], EXPECTED_SOURCE_SHA256)
        self.assertEqual(sha256(SOURCE_PATH), EXPECTED_SOURCE_SHA256)

    def test_06_gold_set_sha_is_unchanged(self) -> None:
        self.assertEqual(self.manifest["gold_set_sha256"], EXPECTED_GOLD_SHA256)
        self.assertEqual(sha256(GOLD_PATH), EXPECTED_GOLD_SHA256)

    def test_07_postfix_results_have_exactly_30_questions(self) -> None:
        self.assertEqual(self.manifest["evaluation_question_count"], 30)
        self.assertEqual(len(self.result_rows), 30)
        self.assertEqual(len({row["eval_id"] for row in self.result_rows}), 30)

    def test_08_question_ids_match_frozen_gold_set(self) -> None:
        gold_ids = [row["eval_id"] for row in self.gold_rows]
        result_ids = [row["eval_id"] for row in self.result_rows]
        self.assertEqual(result_ids, gold_ids)
        self.assertEqual(result_ids, self.manifest["evaluation_question_ids"])

    def test_09_required_scores_are_present_and_not_missing_score(self) -> None:
        for row in self.result_rows:
            self.assertNotIn("MISSING_SCORE", row.values(), row["eval_id"])
            for field in REQUIRED_SCORE_FIELDS:
                self.assertIn(field, row)
                self.assertNotEqual(row[field].strip(), "", f"{row['eval_id']}:{field}")
                float(row[field])

    def test_10_search_failure_count_matches_detail_rows(self) -> None:
        failure_ids = [
            row["eval_id"]
            for row in self.result_rows
            if row.get("search_error", "").strip()
        ]
        self.assertEqual(len(failure_ids), self.manifest["search_failure_count"])
        self.assertEqual(failure_ids, self.manifest["search_failure_ids"])

    def test_11_rate_metrics_are_between_zero_and_one(self) -> None:
        rate_fields = [
            "document_hit_at_1",
            "document_hit_at_3",
            "document_hit_at_5",
            "mrr",
            "element_coverage_at_1",
            "element_coverage_at_3",
            "element_coverage_at_5",
            "metadata_completeness",
        ]
        for metric_set_name in ("baseline_metrics", "postfix_metrics"):
            metric_set = self.manifest[metric_set_name]
            for field in rate_fields:
                self.assertGreaterEqual(float(metric_set[field]), 0.0, field)
                self.assertLessEqual(float(metric_set[field]), 1.0, field)

    def test_12_all_new_result_file_hashes_match_manifest(self) -> None:
        recorded = self.manifest["postfix_result_files_sha256"]
        self.assertTrue(REQUIRED_POSTFIX_FILES.issubset(recorded))
        for name, expected in recorded.items():
            self.assertEqual(sha256(POSTFIX_DIR / name), expected, name)
        self.assertEqual(
            sha256(POSTFIX_DIR / "run_postfix_rag_retrieval_evaluation.py"),
            self.manifest["evaluation_script_sha256"],
        )

    def test_13_postfix_outputs_are_in_separate_folder(self) -> None:
        self.assertEqual(self.manifest["baseline_folder"], BASELINE_DIR.name)
        self.assertEqual(self.manifest["postfix_folder"], POSTFIX_DIR.name)
        actual = {path.name for path in POSTFIX_DIR.iterdir() if path.is_file()}
        self.assertTrue(REQUIRED_POSTFIX_FILES.issubset(actual))
        self.assertTrue(all(not (BASELINE_DIR / name).exists() for name in REQUIRED_POSTFIX_FILES))

    def test_14_evaluation_type_is_postfix_regression(self) -> None:
        self.assertEqual(
            self.manifest["evaluation_type"],
            "post-fix regression evaluation",
        )
        self.assertTrue(self.manifest["same_30_questions_reused"])

    def test_15_not_final_lock_notice_exists(self) -> None:
        notice = self.manifest["not_final_lock_notice"]
        self.assertIn("최종 잠금 평가", notice)
        self.assertIn("아니다", notice)
        self.assertNotIn("최종 성능", notice)
        self.assertIn("법적 정확도", self.manifest["legal_accuracy_notice"])

    def test_16_general_query_evaluation_files_match_manifest_hashes(self) -> None:
        recorded = self.general_query_manifest["result_files_sha256"]
        self.assertTrue(set(recorded).issubset(REQUIRED_GENERAL_QUERY_FILES))
        for name, expected in recorded.items():
            self.assertEqual(sha256(GENERAL_QUERY_DIR / name), expected, name)
        self.assertEqual(
            sha256(GENERAL_QUERY_DIR / "run_general_query_understanding_evaluation.py"),
            self.general_query_manifest["evaluation_script_sha256"],
        )

    def test_17_general_query_counts_and_safety_targets_are_versioned(self) -> None:
        self.assertEqual(self.general_query_manifest["development_question_count"], 72)
        self.assertEqual(self.general_query_manifest["holdout_question_count"], 24)
        for split in ("development_metrics", "holdout_metrics"):
            metrics = self.general_query_manifest[split]
            self.assertEqual(metrics["major_safety_misclassification_count"], 0)
            self.assertGreaterEqual(metrics["primary_domain_accuracy"], 0.93)
            self.assertGreaterEqual(metrics["work_stage_accuracy"], 0.90)
            self.assertGreaterEqual(metrics["requested_output_detection_recall"], 0.95)
            self.assertLessEqual(metrics["maximum_search_query_count"], 4)

    def test_18_general_query_retrieval_regression_meets_minimums(self) -> None:
        metrics = self.general_query_manifest["retrieval_regression_metrics_30"]
        self.assertEqual(metrics["search_failure_count"], 0)
        self.assertGreaterEqual(round(metrics["document_hit_at_3"], 4), 0.5333)
        self.assertGreaterEqual(round(metrics["document_hit_at_5"], 4), 0.5667)
        self.assertGreaterEqual(round(metrics["mrr"], 4), 0.4639)
        self.assertGreaterEqual(round(metrics["metadata_completeness"], 4), 0.7500)

    def test_19_general_query_evaluation_is_not_final_lock(self) -> None:
        notice = self.general_query_manifest["not_final_lock_notice"]
        self.assertIn("최종 잠금 평가", notice)
        self.assertIn("아니다", notice)
        self.assertIn("교수님 최종 5문항", notice)
        self.assertFalse(self.general_query_manifest["external_api_or_internet_used"])
        self.assertTrue(self.general_query_manifest["original_vector_db_unchanged"])


if __name__ == "__main__":
    unittest.main()
