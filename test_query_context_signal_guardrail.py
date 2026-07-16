from __future__ import annotations

import ast
from pathlib import Path
from typing import Any
import unittest


APP_PATH = Path(__file__).with_name("app.py")


def load_query_guardrail_namespace() -> dict[str, Any]:
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8"), filename=str(APP_PATH))
    namespace: dict[str, Any] = {"Any": Any}

    for node in tree.body:
        if not isinstance(node, (ast.Assign, ast.AnnAssign)):
            continue
        target = node.targets[0] if isinstance(node, ast.Assign) and len(node.targets) == 1 else getattr(node, "target", None)
        value_node = node.value
        if not isinstance(target, ast.Name) or value_node is None:
            continue
        try:
            namespace[target.id] = ast.literal_eval(value_node)
        except (ValueError, TypeError):
            continue

    namespace["QUERY_CORE_STATUS_LABELS"] = {
        namespace["QUERY_CORE_STATUS_SUFFICIENT"]: "질문 핵심 요소 반영 충분",
        namespace["QUERY_CORE_STATUS_NEEDS_SUPPLEMENT"]: "질문 핵심 요소 반영 보완 필요",
    }
    namespace["INTENT_SEARCH_EXPANSIONS"] = {
        namespace["BLASTING_MISFIRE_INTENT"]: [
            "불발공", "발파 후 점검", "출입통제", "폭약", "뇌관",
        ],
        namespace["ROOF_FALL_INTENT"]: [
            "천반 점검", "지보", "부석 제거", "균열 확인", "작업중지",
        ],
    }
    intent_values = {
        value
        for name, value in namespace.items()
        if name.endswith("_INTENT") and isinstance(value, str)
    }
    namespace["QUESTION_INTENT_KEYWORDS"] = {
        intent: [] for intent in intent_values
    }

    function_names = {
        "clean_text",
        "normalize_query_question",
        "contains_any_keyword",
        "extract_query_context_signals",
        "detect_question_intent",
        "detect_ppe_item",
        "expand_search_query",
        "build_priority_actions",
        "build_contextual_priority_actions",
        "build_contextual_check_items",
        "_all_keyword_groups_present",
        "_sequence_keywords_present",
        "assess_query_core_element_coverage",
        "build_query_context_supplement",
        "ensure_query_core_elements",
    }
    selected_nodes = [
        node
        for node in tree.body
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        and node.name in function_names
    ]
    exec(compile(ast.Module(selected_nodes, type_ignores=[]), str(APP_PATH), "exec"), namespace)
    return namespace


class QueryContextSignalGuardrailTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = load_query_guardrail_namespace()

    def test_01_rainy_pre_blasting_is_not_fixed_to_misfire_response(self) -> None:
        question = "우천 시 발파 작업 전에 위험 사항을 확인하고 대처방안을 수립하려면?"
        context = self.app["extract_query_context_signals"](question)
        expanded = self.app["expand_search_query"](
            question,
            self.app["BLASTING_MISFIRE_INTENT"],
        )
        self.assertEqual(context["detail_signal"], "rainy_pre_blasting")
        self.assertEqual(context["work_stage"], "발파 작업 전")
        self.assertNotIn("불발공", expanded)
        self.assertNotIn("발파 후 점검", expanded)

    def test_02_rain_blasting_search_and_answer_keep_core_concepts(self) -> None:
        question = "강우 중 발파 전에 무엇을 점검해야 하나?"
        context = self.app["extract_query_context_signals"](question)
        expanded = self.app["expand_search_query"](
            question,
            self.app["BLASTING_MISFIRE_INTENT"],
        )
        actions = " ".join(
            self.app["build_contextual_priority_actions"](
                context,
                self.app["BLASTING_QUESTION_TYPE"],
            )
        )
        combined = expanded + " " + actions
        for concept in ("우천", "낙뢰", "발파공 침수", "화약류", "습윤", "작업을 연기", "중지"):
            self.assertIn(concept, combined)
        coverage = self.app["assess_query_core_element_coverage"](question, actions)
        self.assertEqual(coverage["status"], self.app["QUERY_CORE_STATUS_SUFFICIENT"])

    def test_03_limited_typo_normalization_is_internal(self) -> None:
        original = "대처방안을 수입하려는 광산안전전관리자 질문"
        normalized = self.app["normalize_query_question"](original)
        self.assertEqual(original, "대처방안을 수입하려는 광산안전전관리자 질문")
        self.assertIn("대처방안을 수립", normalized)
        self.assertIn("광산안전관리자", normalized)

    def test_04_roof_water_signal_keeps_roof_support_and_drainage(self) -> None:
        question = "갱내 천반에서 유출수의 양이 증가하면 어떤 조치를 해야 하나?"
        context = self.app["extract_query_context_signals"](question)
        expanded = self.app["expand_search_query"](
            question,
            self.app["ROOF_FALL_INTENT"],
        )
        self.assertEqual(context["detail_signal"], "roof_water_increase")
        for concept in ("유출수", "천반", "지보재 변형", "배수"):
            self.assertIn(concept, expanded)

    def test_05_roof_water_action_order_is_preserved(self) -> None:
        question = "천반 출수가 급증할 때 광산안전관리자의 조치 순서는?"
        context = self.app["extract_query_context_signals"](question)
        actions = self.app["build_contextual_priority_actions"](context, "낙반/붕락")
        answer = " ".join(actions)
        markers = ["작업을 즉시 중지", "대피", "출입을 통제", "보고", "점검", "개선조치", "작업 재개", "기록"]
        positions = [answer.index(marker) for marker in markers]
        self.assertEqual(positions, sorted(positions))
        coverage = self.app["assess_query_core_element_coverage"](question, answer)
        self.assertTrue(coverage["ordered"])
        self.assertEqual(coverage["status"], self.app["QUERY_CORE_STATUS_SUFFICIENT"])

    def test_06_general_roof_and_roof_water_do_not_return_same_fixed_actions(self) -> None:
        general_question = "갱내 천반과 부석 상태를 점검하려면?"
        water_question = "갱내 천반 유출수가 늘어날 때 조치 순서는?"
        general_context = self.app["extract_query_context_signals"](general_question)
        water_context = self.app["extract_query_context_signals"](water_question)
        general_actions = self.app["build_contextual_priority_actions"](general_context, "낙반/붕락")
        water_actions = self.app["build_contextual_priority_actions"](water_context, "낙반/붕락")
        self.assertEqual(general_context["detail_signal"], "")
        self.assertEqual(water_context["detail_signal"], "roof_water_increase")
        self.assertNotEqual(general_actions, water_actions)

    def test_07_explicit_post_blasting_misfire_keeps_existing_response(self) -> None:
        question = "발파 작업 후 불발공이 의심되면 어떻게 해야 해?"
        context = self.app["extract_query_context_signals"](question)
        actions = self.app["build_contextual_priority_actions"](
            context,
            self.app["BLASTING_QUESTION_TYPE"],
        )
        self.assertEqual(context["detail_signal"], "post_blasting_misfire")
        self.assertTrue(any("임의로 접근" in item for item in actions))
        self.assertTrue(any("발파 책임자" in item for item in actions))

    def test_08_out_of_scope_block_is_preserved(self) -> None:
        intent = self.app["detect_question_intent"]("오늘 저녁 메뉴와 맛집을 추천해 줘")
        self.assertEqual(intent, self.app["OUT_OF_SCOPE_INTENT"])


if __name__ == "__main__":
    unittest.main()
