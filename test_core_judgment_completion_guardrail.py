from __future__ import annotations

import ast
from pathlib import Path
from typing import Any
import unittest


APP_PATH = Path(__file__).with_name("app.py")


def load_completion_guardrail_namespace() -> dict[str, Any]:
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

    function_names = {
        "clean_text",
        "generated_answer_incomplete_reason",
        "is_generated_answer_incomplete",
        "required_headings_for_prompt_kind",
        "execute_gemini_with_completion_guardrail",
        "generate_gemini_answer",
        "extract_markdown_section",
        "strip_markdown_prefix",
        "extract_complete_action_items",
        "build_dashboard_summary",
    }
    selected_nodes = [
        node
        for node in tree.body
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        and node.name in function_names
    ]
    exec(compile(ast.Module(selected_nodes, type_ignores=[]), str(APP_PATH), "exec"), namespace)
    return namespace


def success_result(answer: str, *, finish_reason: str = "STOP") -> dict[str, Any]:
    return {
        "called": True,
        "success": True,
        "status": "success",
        "message": "ok",
        "answer": answer,
        "finish_reason": finish_reason,
        "attempts": 1,
        "model": "mock-model",
        "error": "",
        "elapsed": 0.01,
    }


def failure_result() -> dict[str, Any]:
    return {
        "called": True,
        "success": False,
        "status": "error",
        "message": "mock API failure",
        "answer": "",
        "finish_reason": "",
        "attempts": 1,
        "model": "mock-model",
        "error": "mock API failure",
        "elapsed": 0.01,
    }


class CoreJudgmentCompletionGuardrailTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = load_completion_guardrail_namespace()

    def test_01_prompt_fragment_is_incomplete(self) -> None:
        self.assertTrue(self.app["is_generated_answer_incomplete"]("제공된 [안정"))

    def test_02_normal_korean_list_is_not_false_positive(self) -> None:
        answer = """### 우선 조치
- 작업을 중지하고 작업자를 안전한 장소로 대피
- 위험구역 출입통제
- 천반과 지보 상태 확인
- 책임자 확인 후 작업 재개 판단"""
        self.assertFalse(self.app["is_generated_answer_incomplete"](answer))

    def test_03_max_tokens_finish_reason_is_incomplete(self) -> None:
        answer = "작업을 중지하고 대피한 뒤 천반과 지보 상태를 확인하고 책임자에게 보고합니다. 위험 제거 후 재개 여부를 판단합니다."
        reason = self.app["generated_answer_incomplete_reason"](answer, "MAX_TOKENS")
        self.assertIn("비정상 종료 사유", reason)

    def test_04_normal_second_generation_is_used_after_incomplete_first(self) -> None:
        complete = "작업을 즉시 중지하고 안전한 장소로 대피합니다. 위험구역 출입을 통제하고 책임자가 현장을 확인한 뒤 조치와 재개 판단을 기록합니다."
        responses = iter([success_result("제공된 [안정"), success_result(complete)])
        calls: list[str] = []

        def mock_execute(prompt: str, model_name: str) -> dict[str, Any]:
            calls.append(prompt)
            return next(responses)

        self.app["execute_gemini_request"] = mock_execute
        result = self.app["execute_gemini_with_completion_guardrail"](
            "primary",
            "retry",
            "mock-model",
            [],
        )
        self.assertTrue(result["success"])
        self.assertEqual(result["answer"], complete)
        self.assertTrue(result["completion_retry_used"])
        self.assertEqual(len(calls), 2)

    def test_05_two_incomplete_generations_return_fallback_signal(self) -> None:
        responses = iter([success_result("제공된 [안정"), success_result("## 핵심 판단:")])
        calls: list[str] = []

        def mock_execute(prompt: str, model_name: str) -> dict[str, Any]:
            calls.append(prompt)
            return next(responses)

        self.app["execute_gemini_request"] = mock_execute
        result = self.app["execute_gemini_with_completion_guardrail"](
            "primary",
            "retry",
            "mock-model",
            [],
        )
        self.assertFalse(result["success"])
        self.assertEqual(result["status"], "incomplete")
        self.assertEqual(result["answer"], "")
        self.assertEqual(len(calls), 2)

    def configure_generate_stubs(self, execute_results: list[dict[str, Any]]) -> list[str]:
        calls: list[str] = []
        responses = iter(execute_results)

        def mock_execute(prompt: str, model_name: str) -> dict[str, Any]:
            calls.append(prompt)
            return next(responses)

        stable_fallback = (
            "## 즉시 판단\n"
            "- 작업을 중지하고 작업자를 대피시킨 뒤 위험구역 출입을 통제합니다.\n\n"
            "## 우선 조치\n"
            "1. 안전한 위치에서 현장을 점검합니다.\n"
            "2. 위험 제거와 책임자 확인 후 작업 재개 여부를 판단하고 기록합니다."
        )
        coverage = {"status": "sufficient", "label": "질문 핵심 요소 반영 충분"}
        self.app.update(
            {
                "execute_gemini_request": mock_execute,
                "detect_question_intent": lambda question: "mock-intent",
                "build_prompt": lambda *args, **kwargs: "primary",
                "build_hybrid_prompt": lambda *args, **kwargs: "primary",
                "build_completion_retry_prompt": lambda *args, **kwargs: "retry",
                "generate_local_fallback_answer": lambda *args, **kwargs: stable_fallback,
                "ensure_query_core_elements": lambda question, answer: (answer, coverage),
                "build_kras_risk_assessment_section": lambda *args, **kwargs: "### KRAS식 위험성평가 기록 초안\n- mock",
                "build_legacy_gemini_status": lambda result, fallback_used: {
                    **result,
                    "mode": "fallback" if fallback_used else "gemini",
                    "gemini_status": "실패" if fallback_used else "성공",
                    "fallback_used": fallback_used,
                    "used_fallback": fallback_used,
                },
            }
        )
        self.stable_fallback = stable_fallback
        return calls

    def test_06_two_incomplete_answers_use_nonempty_stable_fallback(self) -> None:
        calls = self.configure_generate_stubs(
            [success_result("제공된 [안정"), success_result("## 핵심 판단:")]
        )
        answer, status = self.app["generate_gemini_answer"](
            "천반 유출수 증가 시 조치 순서는?",
            [],
            "mock-model",
            prompt_kind="hybrid",
        )
        self.assertEqual(answer, self.stable_fallback)
        self.assertNotIn("제공된 [안정", answer)
        self.assertTrue(status["fallback_used"])
        self.assertEqual(len(calls), 2)

    def test_07_existing_api_failure_fallback_is_preserved(self) -> None:
        calls = self.configure_generate_stubs([failure_result()])
        answer, status = self.app["generate_gemini_answer"](
            "발파 작업 후 불발이 의심되면?",
            [],
            "mock-model",
            prompt_kind="hybrid",
        )
        self.assertEqual(answer, self.stable_fallback)
        self.assertTrue(status["fallback_used"])
        self.assertEqual(len(calls), 1)

    def test_08_core_judgment_summary_is_complete_one_or_two_sentences(self) -> None:
        self.app.update(
            {
                "build_immediate_judgment": lambda situation: "안정형 핵심판단입니다.",
                "build_priority_actions": lambda situation: ["작업을 중지합니다."],
                "build_check_items": lambda situation: ["책임자가 확인합니다."],
            }
        )
        core_answer = """## 핵심 판단
천반 유출수 증가를 이상징후로 보고 작업을 중지해야 합니다. 위험 제거와 책임자 확인 전에는 작업을 재개하지 않습니다.

## 조치별 설명
1. 작업중지와 대피를 실시합니다.
2. 위험구역 출입을 통제합니다."""
        summary, actions = self.app["build_dashboard_summary"](core_answer, "낙반/붕락")
        self.assertFalse(self.app["is_generated_answer_incomplete"](summary))
        self.assertLessEqual(summary.count("."), 2)
        self.assertTrue(actions)

    def test_09_incomplete_regeneration_never_exceeds_two_calls(self) -> None:
        calls = self.configure_generate_stubs(
            [success_result("제공된 [안정"), success_result("제공된 [안정")]
        )
        self.app["generate_gemini_answer"](
            "우천 시 발파 전 점검은?",
            [],
            "mock-model",
            prompt_kind="hybrid",
        )
        self.assertEqual(len(calls), 2)


if __name__ == "__main__":
    unittest.main()
