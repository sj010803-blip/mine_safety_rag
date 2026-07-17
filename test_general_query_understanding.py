from __future__ import annotations

import ast
from pathlib import Path
import unittest

from query_understanding import (
    DOMAIN_BLASTING,
    DOMAIN_ELECTRICAL,
    DOMAIN_RISK,
    DOMAIN_VENTILATION,
    OUT_OF_SCOPE_DOMAIN,
    analyze_query,
)


MODULE_PATH = Path(__file__).with_name("query_understanding.py")


class GeneralQueryUnderstandingTests(unittest.TestCase):
    def test_01_complex_risk_keeps_primary_and_secondary(self) -> None:
        result = analyze_query("정전으로 환기팬이 멈췄고 가스 측정값도 오르면 무엇부터 해야 해?")
        self.assertEqual(result.primary_domain, DOMAIN_VENTILATION)
        self.assertIn(DOMAIN_ELECTRICAL, result.secondary_domains)
        self.assertIn("환기 정지", result.hazard_signals)
        self.assertIn("전원 상실", result.hazard_signals)
        self.assertIn("가스 상승", result.hazard_signals)

    def test_02_pre_during_post_stages_are_distinct(self) -> None:
        self.assertEqual(analyze_query("발파 작업 전에 점검할 것은?").work_stage, "작업 전")
        self.assertEqual(analyze_query("컨베이어 돌고 있는데 돌을 빼도 돼?").work_stage, "작업 중")
        self.assertEqual(analyze_query("발파 후 불발이 의심되면?").work_stage, "작업 후")

    def test_03_actor_detection(self) -> None:
        self.assertEqual(
            analyze_query("광산안전전관리자가 작업자에게 무엇을 지시해야 하나?").actor,
            "광산안전관리자",
        )
        self.assertEqual(analyze_query("발파책임자가 확인할 것은?").actor, "발파책임자")
        self.assertEqual(analyze_query("구조자가 먼저 확인할 것은?").actor, "구조·응급 대응자")

    def test_04_multiple_requested_outputs_are_preserved(self) -> None:
        result = analyze_query(
            "환기 정지 때 무엇부터 하고 재입갱 기준과 기록, 법적 근거와 사고사례도 알려줘"
        )
        for output in (
            "조치 순서",
            "작업 재개 조건",
            "기록관리",
            "공식 법령·지침 근거",
            "공식 사고사례",
        ):
            self.assertIn(output, result.requested_outputs)

    def test_05_limited_typo_normalization_does_not_change_original(self) -> None:
        original = "광산안전전관리자가 대처방안 수입을 확인한다"
        result = analyze_query(original)
        self.assertEqual(result.original_query, original)
        self.assertIn("광산안전관리자", result.normalized_query)
        self.assertIn("대처방안 수립", result.normalized_query)

    def test_06_ambiguous_water_does_not_assume_roof_or_blast_hole(self) -> None:
        result = analyze_query("물이 갑자기 많이 나오는데 어디서 나오는지 몰라. 계속 작업해도 돼?")
        self.assertEqual(result.primary_domain, DOMAIN_RISK)
        self.assertNotIn("천반", result.hazard_signals)
        self.assertNotIn("발파공 침수", result.hazard_signals)
        self.assertTrue(result.clarification_question)
        self.assertIn("위치", result.clarification_question)

    def test_07_clear_question_is_not_needlessly_clarified(self) -> None:
        result = analyze_query("천반 출수가 늘고 지보재가 휘면 작업자를 어디로 대피시켜야 해?")
        self.assertEqual(result.ambiguity_level, "low")
        self.assertEqual(result.clarification_question, "")

    def test_08_out_of_scope_is_preserved(self) -> None:
        result = analyze_query("오늘 저녁 메뉴와 맛집을 추천해 줘")
        self.assertFalse(result.in_scope)
        self.assertEqual(result.primary_domain, OUT_OF_SCOPE_DOMAIN)
        self.assertEqual(result.requested_outputs, [])

    def test_09_search_plan_contains_original_and_never_exceeds_four(self) -> None:
        question = "천반에서 물이 늘고 지보가 휘는데 조치 순서와 사례도 알려줘"
        result = analyze_query(question)
        self.assertEqual(result.search_queries[0], question)
        self.assertLessEqual(len(result.search_queries), 4)
        self.assertEqual(len(result.search_queries), len(set(result.search_queries)))
        self.assertTrue(result.case_search_query)

    def test_10_no_long_exact_question_comparison_hardcoding(self) -> None:
        tree = ast.parse(MODULE_PATH.read_text(encoding="utf-8"))
        long_exact_comparisons = []
        for node in ast.walk(tree):
            if not isinstance(node, ast.Compare):
                continue
            operands = [node.left, *node.comparators]
            long_literals = [
                item.value
                for item in operands
                if isinstance(item, ast.Constant)
                and isinstance(item.value, str)
                and len(item.value) >= 30
            ]
            if long_literals and any(isinstance(op, (ast.Eq, ast.NotEq)) for op in node.ops):
                long_exact_comparisons.extend(long_literals)
        self.assertEqual(long_exact_comparisons, [])
        source = MODULE_PATH.read_text(encoding="utf-8")
        self.assertNotIn("갱내 천반에서 유출수의 양이 증가하는 경우", source)
        self.assertNotIn("우천 시 발파 작업 전에 발생할 수 있는 위험 사항", source)


if __name__ == "__main__":
    unittest.main()
