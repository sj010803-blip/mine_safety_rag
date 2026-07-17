from __future__ import annotations

import ast
from pathlib import Path
from typing import Any
import unittest

from answer_contract import build_complete_core_judgment, official_case_status_message
from query_understanding import analyze_query


ROOT = Path(__file__).resolve().parent
APP_PATH = ROOT / "app.py"


def load_context_functions() -> dict[str, Any]:
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8"), filename=str(APP_PATH))
    namespace: dict[str, Any] = {
        "Any": Any,
        "analyze_query": analyze_query,
        "build_complete_core_judgment": build_complete_core_judgment,
    }
    for node in tree.body:
        if not isinstance(node, (ast.Assign, ast.AnnAssign)):
            continue
        target = node.targets[0] if isinstance(node, ast.Assign) and len(node.targets) == 1 else getattr(node, "target", None)
        if not isinstance(target, ast.Name) or node.value is None:
            continue
        try:
            namespace[target.id] = ast.literal_eval(node.value)
        except (ValueError, TypeError):
            pass
    selected = [
        node for node in tree.body
        if isinstance(node, ast.FunctionDef)
        and node.name in {
            "clean_text",
            "normalize_query_question",
            "contains_any_keyword",
            "extract_query_context_signals",
        }
    ]
    exec(compile(ast.Module(selected, type_ignores=[]), str(APP_PATH), "exec"), namespace)
    return namespace


class AppQueryUnderstandingIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.namespace = load_context_functions()
        cls.source = APP_PATH.read_text(encoding="utf-8")
        cls.tree = ast.parse(cls.source)

    def context(self, question: str) -> dict[str, Any]:
        return self.namespace["extract_query_context_signals"](question)

    def test_01_legacy_type_and_new_details_are_both_kept(self) -> None:
        context = self.context("비 그친 뒤 발파공에 물이 찬 것 같은데 장약 전에 뭘 확인해?")
        self.assertEqual(context["detail_signal"], "rainy_pre_blasting")
        self.assertIn("발파", context["major_risk_type"])
        self.assertIn("우천", context["hazard_signals"])
        self.assertIn("발파공 침수", context["hazard_signals"])
        self.assertTrue(context["display_label"])

    def test_02_previous_three_questions_keep_distinct_context(self) -> None:
        rain = self.context("우천 시 발파 작업 전에 대처방안을 수입하려면?")
        roof = self.context("천반 유출수가 늘면 광산안전전관리자가 어떤 순서로 조치해야 하나?")
        misfire = self.context("발파 작업 후 불발공이 의심되면 어떻게 해야 해?")
        self.assertEqual(rain["detail_signal"], "rainy_pre_blasting")
        self.assertEqual(roof["detail_signal"], "roof_water_increase")
        self.assertEqual(misfire["detail_signal"], "post_blasting_misfire")
        self.assertNotEqual(rain["hazard_signals"], misfire["hazard_signals"])

    def test_03_new_paraphrase_uses_same_general_hazard_context(self) -> None:
        first = analyze_query("정전으로 환기팬이 멈추고 가스가 오르면 무엇부터 해야 해?")
        second = analyze_query("송풍이 끊긴 뒤 검지기 값도 높아져. 사람부터 빼야 해?")
        self.assertEqual(first.primary_domain, second.primary_domain)
        self.assertIn("환기 정지", first.hazard_signals)
        self.assertIn("환기 정지", second.hazard_signals)

    def test_04_official_evidence_guardrail_remains_in_flow(self) -> None:
        self.assertIn("assess_rag_evidence_sufficiency", self.source)
        self.assertIn("enforce_rag_evidence_answer_guardrail", self.source)
        self.assertIn("실제 문서명·chunk_id", APP_PATH.with_name("answer_contract.py").read_text(encoding="utf-8"))

    def test_05_existing_gemini_completion_fallback_remains(self) -> None:
        self.assertIn("execute_gemini_with_completion_guardrail", self.source)
        self.assertIn("generate_local_fallback_answer", self.source)
        self.assertIn("build_rule_based_fallback", self.source)
        self.assertIn("completion_retry_used", self.source)

    def test_06_out_of_scope_returns_before_rag_or_external_cases(self) -> None:
        run_node = next(
            node for node in self.tree.body
            if isinstance(node, ast.FunctionDef) and node.name == "run_rag_flow"
        )
        search_lines = [
            node.lineno for node in ast.walk(run_node)
            if isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id in {"search_vector_db", "search_official_siren_cases", "get_live_reference_cases"}
        ]
        out_scope_returns = [
            node.lineno for node in ast.walk(run_node)
            if isinstance(node, ast.Return)
            and isinstance(node.value, ast.Tuple)
            and any(isinstance(item, ast.Name) and item.id == "out_status" for item in node.value.elts)
        ]
        self.assertTrue(search_lines and out_scope_returns)
        self.assertLess(min(out_scope_returns), min(search_lines))

    def test_07_no_official_case_does_not_display_unrelated_case(self) -> None:
        understanding = analyze_query("컨베이어 끼임 실제 사고사례도 알려줘")
        message = official_case_status_message(understanding, [])
        self.assertIn("무관한 사례는 표시하지 않습니다", message)

    def test_08_three_answer_modes_remain_defined(self) -> None:
        for marker in ("STABLE_MODE", "WORKER_EASY_MODE", "HYBRID_MODE"):
            self.assertIn(marker, self.source)
        self.assertIn("if answer_mode == STABLE_MODE", self.source)
        self.assertIn("if answer_mode == HYBRID_MODE", self.source)


if __name__ == "__main__":
    unittest.main()
