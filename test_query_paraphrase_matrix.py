from __future__ import annotations

import ast
from collections import Counter
import json
from pathlib import Path
import unittest

from query_understanding import OUT_OF_SCOPE_DOMAIN, SUPPORTED_DOMAINS, analyze_query


ROOT = Path(__file__).resolve().parent
EVAL_DIR = ROOT / "28_general_query_understanding_evaluation"
DEVELOPMENT_PATH = EVAL_DIR / "development_question_matrix.jsonl"
HOLDOUT_PATH = EVAL_DIR / "holdout_question_matrix.jsonl"
MODULE_PATH = ROOT / "query_understanding.py"
REQUIRED_FIELDS = {
    "case_id",
    "split",
    "question",
    "expected_primary_domain",
    "expected_secondary_domains",
    "required_signals",
    "expected_stage",
    "required_requested_outputs",
    "should_clarify",
    "forbidden_dominant_topics",
    "note",
}


def read_jsonl(path: Path) -> list[dict]:
    return [
        json.loads(line)
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]


def metrics(rows: list[dict]) -> dict[str, float | int]:
    analyzed = [analyze_query(row["question"]) for row in rows]
    required_signal_count = sum(len(row["required_signals"]) for row in rows)
    requested_count = sum(len(row["required_requested_outputs"]) for row in rows)
    secondary_count = sum(len(row["expected_secondary_domains"]) for row in rows)
    return {
        "primary_domain_accuracy": sum(
            result.primary_domain == row["expected_primary_domain"]
            for result, row in zip(analyzed, rows)
        ) / len(rows),
        "secondary_domain_recall": (
            sum(
                len(set(row["expected_secondary_domains"]) & set(result.secondary_domains))
                for result, row in zip(analyzed, rows)
            ) / secondary_count
            if secondary_count else 1.0
        ),
        "hazard_signal_coverage": sum(
            len(set(row["required_signals"]) & set(result.hazard_signals))
            for result, row in zip(analyzed, rows)
        ) / required_signal_count,
        "work_stage_accuracy": sum(
            result.work_stage == row["expected_stage"]
            for result, row in zip(analyzed, rows)
        ) / len(rows),
        "requested_output_recall": sum(
            len(set(row["required_requested_outputs"]) & set(result.requested_outputs))
            for result, row in zip(analyzed, rows)
        ) / requested_count,
        "ambiguity_accuracy": sum(
            bool(result.clarification_question) == bool(row["should_clarify"])
            for result, row in zip(analyzed, rows)
        ) / len(rows),
        "major_misclassification_count": sum(
            result.primary_domain != row["expected_primary_domain"]
            for result, row in zip(analyzed, rows)
        ),
        "average_search_query_count": sum(len(result.search_queries) for result in analyzed) / len(rows),
        "max_search_query_count": max(len(result.search_queries) for result in analyzed),
    }


class QueryParaphraseMatrixTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.development = read_jsonl(DEVELOPMENT_PATH)
        cls.holdout = read_jsonl(HOLDOUT_PATH)
        cls.all_rows = cls.development + cls.holdout

    def test_01_minimum_split_counts(self) -> None:
        self.assertGreaterEqual(len(self.development), 72)
        self.assertGreaterEqual(len(self.holdout), 24)

    def test_02_at_least_twelve_domains_are_represented(self) -> None:
        counts = Counter(row["expected_primary_domain"] for row in self.all_rows)
        self.assertTrue(set(SUPPORTED_DOMAINS).issubset(counts))
        self.assertGreaterEqual(min(counts[domain] for domain in SUPPORTED_DOMAINS), 6)

    def test_03_case_ids_are_unique(self) -> None:
        ids = [row["case_id"] for row in self.all_rows]
        self.assertEqual(len(ids), len(set(ids)))

    def test_04_question_sentences_are_unique(self) -> None:
        questions = [row["question"] for row in self.all_rows]
        self.assertEqual(len(questions), len(set(questions)))

    def test_05_gold_fields_are_complete(self) -> None:
        for row in self.all_rows:
            self.assertTrue(REQUIRED_FIELDS.issubset(row), row.get("case_id"))
            self.assertTrue(row["question"].strip())
            self.assertIn(row["split"], {"development", "holdout"})
            self.assertTrue(row["required_signals"])
            self.assertTrue(row["required_requested_outputs"])

    def test_06_development_performance_targets(self) -> None:
        score = metrics(self.development)
        self.assertGreaterEqual(score["primary_domain_accuracy"], 0.93)
        self.assertGreaterEqual(score["work_stage_accuracy"], 0.90)
        self.assertGreaterEqual(score["requested_output_recall"], 0.95)
        self.assertEqual(score["major_misclassification_count"], 0)
        self.assertLessEqual(score["max_search_query_count"], 4)

    def test_07_holdout_has_no_major_safety_misclassification(self) -> None:
        score = metrics(self.holdout)
        self.assertGreaterEqual(score["primary_domain_accuracy"], 0.93)
        self.assertGreaterEqual(score["work_stage_accuracy"], 0.90)
        self.assertGreaterEqual(score["requested_output_recall"], 0.95)
        self.assertEqual(score["major_misclassification_count"], 0)
        self.assertLessEqual(score["max_search_query_count"], 4)

    def test_08_out_of_scope_block_accuracy(self) -> None:
        questions = (
            "오늘 저녁 메뉴 추천해 줘",
            "주식 종목을 골라 줘",
            "이번 주말 여행지는 어디가 좋아?",
            "야구 경기 결과 알려줘",
        )
        results = [analyze_query(question) for question in questions]
        self.assertTrue(all(result.primary_domain == OUT_OF_SCOPE_DOMAIN for result in results))
        self.assertTrue(all(not result.in_scope for result in results))

    def test_09_no_exact_full_sentence_hardcoding(self) -> None:
        source = MODULE_PATH.read_text(encoding="utf-8")
        for row in self.all_rows:
            self.assertNotIn(row["question"], source)
        tree = ast.parse(source)
        exact_long_literals = []
        for node in ast.walk(tree):
            if isinstance(node, ast.Compare) and any(isinstance(op, ast.Eq) for op in node.ops):
                for value in (node.left, *node.comparators):
                    if isinstance(value, ast.Constant) and isinstance(value.value, str) and len(value.value) >= 30:
                        exact_long_literals.append(value.value)
        self.assertEqual(exact_long_literals, [])

    def test_10_search_query_statistics_are_bounded(self) -> None:
        for rows in (self.development, self.holdout):
            score = metrics(rows)
            self.assertGreaterEqual(score["average_search_query_count"], 1.0)
            self.assertLessEqual(score["max_search_query_count"], 4)


if __name__ == "__main__":
    unittest.main()
