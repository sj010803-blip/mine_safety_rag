"""RAG 공식 근거 검색 상태 안전장치의 비파괴 정적 테스트."""

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
        self.assigned_names = {
            target.id
            for node in ast.walk(self.tree)
            if isinstance(node, (ast.Assign, ast.AnnAssign))
            for target in (
                node.targets
                if isinstance(node, ast.Assign)
                else [node.target]
            )
            if isinstance(target, ast.Name)
        }
        self.string_text = "\n".join(strings_in(self.tree))

    def function(self, name: str) -> ast.FunctionDef | ast.AsyncFunctionDef:
        return self.functions[name]

    def source(self, name: str) -> str:
        node = self.function(name)
        return "\n".join(self.lines[node.lineno - 1 : node.end_lineno])


class RagEvidenceSufficiencyGuardrailTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not APP_PATH.is_file():
            raise AssertionError("app.py 파일이 없습니다.")
        cls.app = StaticApp(APP_PATH)

    def test_01_assessment_structure_and_adjustable_thresholds(self) -> None:
        app = self.app
        self.assertIn("assess_rag_evidence_sufficiency", app.functions)
        required_constants = {
            "EVIDENCE_MIN_OFFICIAL_CHUNKS",
            "EVIDENCE_MIN_UNIQUE_DOCUMENTS",
            "EVIDENCE_MIN_NON_EMPTY_CHUNKS",
            "EVIDENCE_MAX_DUPLICATE_RATIO",
            "EVIDENCE_MAX_SINGLE_DOCUMENT_RATIO",
        }
        self.assertTrue(required_constants.issubset(app.assigned_names))
        assessment_strings = set(strings_in(app.function("assess_rag_evidence_sufficiency")))
        required_keys = {
            "status",
            "label",
            "reason",
            "official_chunk_count",
            "unique_document_count",
            "non_empty_chunk_count",
            "duplicate_chunk_count",
            "diagnostic_details",
        }
        self.assertTrue(required_keys.issubset(assessment_strings))

    def test_02_three_korean_statuses_and_explanations(self) -> None:
        for label in ("공식 근거 충분", "공식 근거 보완 필요", "공식 근거 부족"):
            self.assertIn(label, self.app.string_text)
        for explanation in (
            "주요 안전조치를 뒷받침",
            "일부 세부 기준은 추가 확인",
            "법령·수치·작업 재개 조건을 단정하기 어렵습니다",
        ):
            self.assertIn(explanation, self.app.string_text)

    def test_03_not_accuracy_percentage_and_distance_not_used(self) -> None:
        misleading_percentage = re.compile(
            r"(?:정확도|법적\s*신뢰도)\s*(?:점수\s*)?\d+(?:\.\d+)?\s*%"
        )
        self.assertIsNone(misleading_percentage.search(self.app.text))
        assessment = self.app.function("assess_rag_evidence_sufficiency")
        distance_get_calls = []
        for call in calls_in(assessment):
            if dotted_name(call.func) != "result.get" or not call.args:
                continue
            first = call.args[0]
            if isinstance(first, ast.Constant) and first.value == "distance":
                distance_get_calls.append(call)
        self.assertFalse(distance_get_calls, "distance가 충분성 판정 입력으로 사용되고 있습니다.")
        source = self.app.source("assess_rag_evidence_sufficiency")
        self.assertNotRegex(source, r":\.\d+%|:\.\d+%")
        self.assertIn("충분성 판정에는 사용하지 않음", self.app.string_text)

    def test_04_out_of_scope_returns_before_search_and_panel(self) -> None:
        app = self.app
        run_node = app.function("run_rag_flow")
        out_scope_branches = [
            node
            for node in ast.walk(run_node)
            if isinstance(node, ast.If) and "OUT_OF_SCOPE_INTENT" in names_in(node.test)
        ]
        self.assertTrue(out_scope_branches)
        self.assertTrue(any(isinstance(item, ast.Return) for item in ast.walk(out_scope_branches[0])))
        search_line = min(call.lineno for call in calls_in(run_node, "search_vector_db"))
        assessment_line = min(
            call.lineno
            for call in calls_in(run_node, "assess_rag_evidence_sufficiency")
        )
        self.assertLess(out_scope_branches[0].lineno, search_line)
        self.assertLess(search_line, assessment_line)

        render_node = app.function("render_rag_result")
        render_out_scope = [
            node
            for node in ast.walk(render_node)
            if isinstance(node, ast.If) and "OUT_OF_SCOPE_INTENT" in names_in(node.test)
        ]
        panel_line = min(
            call.lineno
            for call in calls_in(render_node, "render_rag_evidence_guardrail")
        )
        self.assertTrue(render_out_scope)
        self.assertTrue(any(isinstance(item, ast.Return) for item in ast.walk(render_out_scope[0])))
        self.assertLess(render_out_scope[0].lineno, panel_line)

    def test_05_one_shared_assessment_for_all_answer_modes(self) -> None:
        app = self.app
        run_node = app.function("run_rag_flow")
        self.assertEqual(
            len(calls_in(run_node, "assess_rag_evidence_sufficiency")),
            1,
        )
        run_strings = set(strings_in(run_node))
        self.assertTrue(
            {"evidence_assessment", "evidence_status", "evidence_label", "evidence_reason"}.issubset(
                run_strings
            )
        )
        gemini_calls = calls_in(run_node, "generate_gemini_answer")
        self.assertEqual(len(gemini_calls), 2)
        for call in gemini_calls:
            keyword_names = {keyword.arg for keyword in call.keywords}
            self.assertIn("evidence_assessment", keyword_names)
        local_calls = calls_in(run_node, "generate_local_fallback_answer")
        shared_local_calls = [
            call
            for call in local_calls
            if "evidence_assessment" in {keyword.arg for keyword in call.keywords}
        ]
        self.assertEqual(len(shared_local_calls), 1)
        self.assertIn(
            "build_evidence_guardrail_prompt_guidance",
            {dotted_name(call.func) for call in calls_in(app.function("build_prompt"))},
        )
        self.assertIn(
            "build_evidence_guardrail_prompt_guidance",
            {dotted_name(call.func) for call in calls_in(app.function("build_hybrid_prompt"))},
        )
        guardrail_calls = calls_in(run_node, "enforce_rag_evidence_answer_guardrail")
        self.assertEqual(len(guardrail_calls), 3)
        for call in guardrail_calls:
            self.assertIn("evidence_assessment", names_in(call))

    def test_06_news_and_cases_are_not_official_legal_evidence(self) -> None:
        guidance = "\n".join(
            strings_in(self.app.function("build_evidence_guardrail_prompt_guidance"))
        )
        self.assertIn("뉴스는 공식 법령 판단 근거로 사용하지 마세요", guidance)
        self.assertIn("사례 기반 주의 포인트는 공식 법령 판단 근거로 사용하지 마세요", guidance)
        self.assertIn("RAG 문서명과 chunk_id", guidance)
        self.assertIn("중대재해처벌법 위반을 확정적으로 단정하지 마세요", guidance)

    def test_07_insufficient_evidence_requires_conservative_action(self) -> None:
        text = self.app.string_text
        self.assertIn("검색된 공식 문서만으로는 세부 기준을 충분히 확인하기 어렵습니다", text)
        self.assertIn("담당 안전관리자, 관계기관 또는 전문가", text)
        for action in (
            "즉시 작업중지",
            "위험구역 접근통제",
            "작업자 대피",
            "담당 안전관리자 보고",
            "승인 전 작업 재개 금지",
        ):
            self.assertIn(action, text)
        output_guardrail = self.app.function("enforce_rag_evidence_answer_guardrail")
        output_text = "\n".join(strings_in(output_guardrail))
        self.assertIn("현재 검색 결과에 없는 법령 조항 번호", output_text)
        self.assertIn("개선조치 완료 확인 및 책임자 승인 전 작업 재개 금지", output_text)
        self.assertIn("중대재해처벌법상 쟁점이 될 수 있음", output_text)

    def test_08_admin_diagnostics_are_separate_and_textual(self) -> None:
        app = self.app
        panel = app.function("render_rag_evidence_guardrail")
        panel_calls = {dotted_name(call.func) for call in calls_in(panel)}
        self.assertIn("st.expander", panel_calls)
        self.assertIn("is_admin_mode", names_in(panel))
        panel_text = "\n".join(strings_in(panel))
        for item in (
            "공식 chunk 수",
            "고유 문서 수",
            "비어 있지 않은 chunk 수",
            "중복 결과 수",
            "사용된 판정 기준",
            "검색 점수 해석 방식",
            "판정 사유",
        ):
            self.assertIn(item, panel_text)
        render_calls = {
            dotted_name(call.func)
            for call in calls_in(app.function("render_rag_result"))
        }
        self.assertIn("render_rag_evidence_guardrail", render_calls)

    def test_09_guardrail_does_not_read_or_print_secret_values(self) -> None:
        app = self.app
        forbidden_calls = {
            "os.getenv",
            "os.environ.get",
            "get_gemini_api_key",
        }
        for function_name in (
            "assess_rag_evidence_sufficiency",
            "build_evidence_guardrail_prompt_guidance",
            "render_rag_evidence_guardrail",
        ):
            node = app.function(function_name)
            function_calls = {dotted_name(call.func) for call in calls_in(node)}
            self.assertFalse(forbidden_calls & function_calls)
            self.assertNotIn("st.secrets", app.source(function_name))
            function_text = "\n".join(strings_in(node))
            for secret_name in ("GEMINI_API_KEY", "NAVER_CLIENT_ID", "NAVER_CLIENT_SECRET"):
                self.assertNotIn(secret_name, function_text)

    def test_10_history_and_excel_columns_are_connected(self) -> None:
        app = self.app
        history_strings = set(strings_in(app.function("auto_save_conversation_history")))
        self.assertTrue(
            {"evidence_status", "evidence_label", "evidence_reason"}.issubset(history_strings)
        )
        normalize_text = "\n".join(
            strings_in(app.function("normalize_conversation_history_rows"))
        )
        self.assertIn("공식 근거 검색 상태", normalize_text)
        self.assertIn("공식 근거 판정 사유", normalize_text)
        save_calls = calls_in(app.function("render_rag_result"), "auto_save_conversation_history")
        self.assertEqual(len(save_calls), 1)
        self.assertIn(
            "evidence_assessment",
            {keyword.arg for keyword in save_calls[0].keywords},
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
