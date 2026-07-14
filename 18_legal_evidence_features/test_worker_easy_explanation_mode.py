"""근로자 쉬운 설명 모드의 비파괴 정적 회귀 테스트."""

from __future__ import annotations

import ast
from pathlib import Path
import re
import unittest


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def dotted_name(node: ast.AST) -> str:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        base = dotted_name(node.value)
        return f"{base}.{node.attr}" if base else node.attr
    return ""


def calls_in(node: ast.AST, name: str | None = None) -> list[ast.Call]:
    calls = [item for item in ast.walk(node) if isinstance(item, ast.Call)]
    if name is None:
        return calls
    return [item for item in calls if dotted_name(item.func) == name]


def strings_in(node: ast.AST) -> list[str]:
    return [
        item.value
        for item in ast.walk(node)
        if isinstance(item, ast.Constant) and isinstance(item.value, str)
    ]


def names_in(node: ast.AST) -> set[str]:
    return {item.id for item in ast.walk(node) if isinstance(item, ast.Name)}


class StaticApp:
    def __init__(self, path: Path) -> None:
        self.text = path.read_text(encoding="utf-8-sig")
        self.lines = self.text.splitlines()
        self.tree = ast.parse(self.text, filename=str(path))
        self.functions = {
            node.name: node
            for node in ast.walk(self.tree)
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        }
        self.assignments: dict[str, ast.AST] = {}
        for node in ast.walk(self.tree):
            if not isinstance(node, (ast.Assign, ast.AnnAssign)):
                continue
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            value = node.value
            if value is None:
                continue
            for target in targets:
                if isinstance(target, ast.Name):
                    self.assignments[target.id] = value
        self.string_text = "\n".join(strings_in(self.tree))

    def function(self, name: str) -> ast.FunctionDef | ast.AsyncFunctionDef:
        return self.functions[name]

    def source(self, name: str) -> str:
        node = self.function(name)
        return "\n".join(self.lines[node.lineno - 1 : node.end_lineno])

    def assigned_string(self, name: str) -> str | None:
        node = self.assignments.get(name)
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            return node.value
        return None


class WorkerEasyExplanationModeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not APP_PATH.is_file():
            raise AssertionError("app.py 파일이 없습니다.")
        cls.app = StaticApp(APP_PATH)

    def test_01_worker_easy_display_name_and_selectbox_mapping(self) -> None:
        app = self.app
        self.assertEqual(
            app.assigned_string("WORKER_EASY_MODE_LABEL"),
            "근로자 쉬운 설명 모드",
        )
        self.assertIn("근로자 쉬운 설명 모드", app.string_text)
        selectboxes = calls_in(app.tree, "st.sidebar.selectbox")
        answer_mode_selectboxes = [
            call
            for call in selectboxes
            if call.args
            and isinstance(call.args[0], ast.Constant)
            and call.args[0].value == "답변 생성 방식"
        ]
        self.assertEqual(len(answer_mode_selectboxes), 1)
        format_keywords = [
            keyword
            for keyword in answer_mode_selectboxes[0].keywords
            if keyword.arg == "format_func"
        ]
        self.assertEqual(len(format_keywords), 1)
        self.assertEqual(dotted_name(format_keywords[0].value), "answer_mode_option_label")

    def test_02_prompt_targets_new_and_non_expert_workers(self) -> None:
        prompt_text = "\n".join(strings_in(self.app.function("build_prompt")))
        self.assertRegex(prompt_text, r"신규\s*광산\s*근로자")
        self.assertIn("비전문 작업자", prompt_text)
        self.assertIn("성인 신규 근로자", prompt_text)
        self.assertIn("지나치게 유치하게", prompt_text)

    def test_03_prompt_requires_why_the_hazard_is_dangerous(self) -> None:
        prompt_text = "\n".join(strings_in(self.app.function("build_prompt")))
        self.assertIn("왜 위험한가요?", prompt_text)
        self.assertRegex(prompt_text, r"끼임.*질식.*폭발.*낙하.*감전.*충돌")
        self.assertIn("worker_easy_intent_guidance", self.app.functions)

    def test_04_professional_terms_have_plain_language_direction(self) -> None:
        app = self.app
        self.assertIn("WORKER_EASY_TERM_EXPLANATIONS", app.assignments)
        term_text = "\n".join(strings_in(app.assignments["WORKER_EASY_TERM_EXPLANATIONS"]))
        for term in ("격리조치", "에너지 차단", "산소결핍", "불발공", "작업 재개 승인"):
            self.assertIn(term, term_text)
        self.assertGreaterEqual(term_text.count("→"), 5)
        self.assertIn("WORKER_EASY_TERM_EXPLANATIONS", names_in(app.function("build_prompt")))

    def test_05_virtual_examples_are_limited_and_not_official_cases(self) -> None:
        prompt_text = "\n".join(strings_in(self.app.function("build_prompt")))
        self.assertIn("이해를 돕는 가상 예시", prompt_text)
        self.assertRegex(prompt_text, r"최대\s*2개")
        self.assertIn("실제 공식 사고사례처럼 표현하지 마세요", prompt_text)

    def test_06_unverified_numbers_time_concentration_and_law_are_banned(self) -> None:
        prompt_text = "\n".join(strings_in(self.app.function("build_prompt")))
        self.assertRegex(
            prompt_text,
            r"확인되지 않은\s*수치,\s*시간,\s*농도,\s*법령 조문을 만들지 마세요",
        )
        self.assertIn("RAG 근거 밖의 사실을 추가하지 마세요", prompt_text)
        self.assertIn("모른다고 분명하게", prompt_text)

    def test_07_news_and_cases_are_not_official_evidence(self) -> None:
        prompt_text = "\n".join(strings_in(self.app.function("build_prompt")))
        self.assertIn("사례 기반 주의 포인트", prompt_text)
        self.assertIn("공식 법령 근거가 아니라", prompt_text)
        self.assertIn("뉴스는 공식 근거가 아니므로", prompt_text)
        self.assertIn("법적 판단이나 가상 예시에 사용하지 마세요", prompt_text)

    def test_08_major_accident_law_violation_is_not_declared(self) -> None:
        prompt_text = "\n".join(strings_in(self.app.function("build_prompt")))
        self.assertIn("중대재해처벌법 위반을 확정적으로 단정하지 마세요", prompt_text)
        self.assertIn("처벌을 강조하지 마세요", prompt_text)

    def test_09_worker_prompt_and_hybrid_prompt_are_separate(self) -> None:
        app = self.app
        generator = app.function("generate_gemini_answer")
        generator_calls = {dotted_name(call.func) for call in calls_in(generator)}
        self.assertIn("build_prompt", generator_calls)
        self.assertIn("build_hybrid_prompt", generator_calls)
        worker_text = "\n".join(strings_in(app.function("build_prompt")))
        hybrid_text = "\n".join(strings_in(app.function("build_hybrid_prompt")))
        self.assertIn("### 지금 바로 해야 할 일", worker_text)
        self.assertNotIn("### 지금 바로 해야 할 일", hybrid_text)
        self.assertIn("## 조치별 설명", hybrid_text)
        self.assertNotIn("## 조치별 설명", worker_text)

    def test_10_hybrid_keeps_stable_draft_and_manager_detail(self) -> None:
        hybrid = self.app.function("build_hybrid_prompt")
        hybrid_calls = {dotted_name(call.func) for call in calls_in(hybrid)}
        hybrid_text = "\n".join(strings_in(hybrid))
        self.assertIn("build_context", hybrid_calls)
        self.assertIn("build_evidence_guardrail_prompt_guidance", hybrid_calls)
        self.assertIn("stable_draft", names_in(hybrid))
        self.assertIn("현장관리자와 안전관리자", hybrid_text)
        self.assertIn("점검·통제·보고·기록관리", hybrid_text)
        self.assertIn("법적·관리적 맥락", hybrid_text)

    def test_11_gemini_failure_has_stable_fallback_and_short_notice(self) -> None:
        app = self.app
        generator_calls = {
            dotted_name(call.func)
            for call in calls_in(app.function("generate_gemini_answer"))
        }
        self.assertIn("generate_local_fallback_answer", generator_calls)
        run_text = "\n".join(strings_in(app.function("run_rag_flow")))
        self.assertIn(
            "쉬운 설명을 생성하지 못해 공식 문서 기반 핵심조치로 대신 안내합니다.",
            run_text,
        )
        self.assertIn("Gemini API 호출 실패 → 안정형 fallback", run_text)

    def test_12_out_of_scope_returns_before_worker_generation(self) -> None:
        run_node = self.app.function("run_rag_flow")
        branches = [
            node
            for node in ast.walk(run_node)
            if isinstance(node, ast.If) and "OUT_OF_SCOPE_INTENT" in names_in(node.test)
        ]
        self.assertTrue(branches)
        self.assertTrue(any(isinstance(node, ast.Return) for node in ast.walk(branches[0])))
        search_line = min(call.lineno for call in calls_in(run_node, "search_vector_db"))
        gemini_line = min(call.lineno for call in calls_in(run_node, "generate_gemini_answer"))
        self.assertLess(branches[0].lineno, search_line)
        self.assertLess(search_line, gemini_line)

    def test_13_legacy_mode_values_and_history_are_compatible(self) -> None:
        app = self.app
        self.assertEqual(
            app.assigned_string("GEMINI_MODE"),
            "자연어 설명 모드: 검색 근거 기반 설명형 답변",
        )
        normalize_text = "\n".join(strings_in(app.function("normalize_answer_mode_label")))
        for legacy in ("자연어 설명 모드", "Gemini 모드", "natural", "gemini"):
            self.assertIn(legacy, normalize_text)
        history_node = app.function("normalize_conversation_history_rows")
        self.assertIn(
            "normalize_answer_mode_label",
            {dotted_name(call.func) for call in calls_in(history_node)},
        )
        self.assertIn("답변 모드", strings_in(history_node))
        self.assertIn("answer_mode", strings_in(app.function("auto_save_conversation_history")))

    def test_14_shared_rag_evidence_status_is_reused(self) -> None:
        app = self.app
        run_node = app.function("run_rag_flow")
        self.assertEqual(len(calls_in(run_node, "assess_rag_evidence_sufficiency")), 1)
        for call in calls_in(run_node, "generate_gemini_answer"):
            self.assertIn("evidence_assessment", {keyword.arg for keyword in call.keywords})
        worker_prompt_calls = {
            dotted_name(call.func)
            for call in calls_in(app.function("build_prompt"))
        }
        self.assertIn("build_evidence_guardrail_prompt_guidance", worker_prompt_calls)
        self.assertIn("build_worker_easy_evidence_guidance", worker_prompt_calls)
        self.assertFalse(
            calls_in(app.function("build_worker_easy_evidence_guidance"), "assess_rag_evidence_sufficiency")
        )
        guardrail_calls = calls_in(run_node, "enforce_rag_evidence_answer_guardrail")
        worker_calls = [
            call
            for call in guardrail_calls
            if any(
                keyword.arg == "worker_easy"
                and isinstance(keyword.value, ast.Constant)
                and keyword.value.value is True
                for keyword in call.keywords
            )
        ]
        self.assertEqual(len(worker_calls), 1)
        easy_guardrail_text = "\n".join(
            strings_in(app.function("enforce_rag_evidence_answer_guardrail"))
        )
        self.assertIn("### 왜 위험한가요?", easy_guardrail_text)
        self.assertIn("다시 작업해도 된다고 할 때까지 기다리세요", easy_guardrail_text)

    def test_15_worker_mode_does_not_read_or_print_secret_values(self) -> None:
        app = self.app
        forbidden_calls = {"os.getenv", "os.environ.get", "get_gemini_api_key"}
        for function_name in (
            "worker_easy_intent_guidance",
            "build_worker_easy_evidence_guidance",
            "build_prompt",
            "normalize_answer_mode_label",
        ):
            node = app.function(function_name)
            function_calls = {dotted_name(call.func) for call in calls_in(node)}
            self.assertFalse(forbidden_calls & function_calls)
            self.assertNotIn("st.secrets", app.source(function_name))
            function_text = "\n".join(strings_in(node))
            self.assertIsNone(
                re.search(r"(?:api[_ -]?key|client[_ -]?secret)\s*[:=]\s*['\"][^'\"]+", function_text, re.I)
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
