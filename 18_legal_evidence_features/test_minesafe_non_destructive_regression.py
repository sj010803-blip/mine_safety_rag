"""MineSafe AI 비파괴 정적 회귀 테스트.

이 파일은 app.py를 import하거나 실행하지 않는다. app.py를 UTF-8 텍스트로 읽고
ast.parse()로 구조만 분석한다. Streamlit, 외부 API, Vector DB에도 연결하지 않는다.
"""

from __future__ import annotations

import ast
import hashlib
from pathlib import Path
import re
import unittest


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def dotted_name(node: ast.AST) -> str:
    """Name/Attribute 노드를 점으로 연결한 이름으로 바꾼다."""
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        base = dotted_name(node.value)
        return f"{base}.{node.attr}" if base else node.attr
    return ""


def target_names(node: ast.AST) -> set[str]:
    """대입 대상에서 변수 이름만 안전하게 수집한다."""
    names: set[str] = set()
    if isinstance(node, ast.Name):
        names.add(node.id)
    elif isinstance(node, (ast.Tuple, ast.List)):
        for item in node.elts:
            names.update(target_names(item))
    return names


def names_in(node: ast.AST) -> set[str]:
    return {item.id for item in ast.walk(node) if isinstance(item, ast.Name)}


def calls_in(node: ast.AST) -> set[str]:
    return {
        dotted_name(item.func)
        for item in ast.walk(node)
        if isinstance(item, ast.Call) and dotted_name(item.func)
    }


def strings_in(node: ast.AST) -> list[str]:
    return [
        item.value
        for item in ast.walk(node)
        if isinstance(item, ast.Constant) and isinstance(item.value, str)
    ]


def has_return_in(node: ast.AST) -> bool:
    return any(isinstance(item, ast.Return) for item in ast.walk(node))


def contains_empty_list_return(node: ast.AST) -> bool:
    for item in ast.walk(node):
        if not isinstance(item, ast.Return):
            continue
        if isinstance(item.value, ast.List) and not item.value.elts:
            return True
    return False


class StaticAppModel:
    """app.py의 실행 없는 정적 표현."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.raw = path.read_bytes()
        self.text = self.raw.decode("utf-8-sig")
        self.tree = ast.parse(self.text, filename=str(path))
        self.size = len(self.raw)
        self.sha256 = hashlib.sha256(self.raw).hexdigest().upper()
        self.mtime_ns = path.stat().st_mtime_ns

        self.functions: dict[str, ast.FunctionDef | ast.AsyncFunctionDef] = {}
        self.imports: set[str] = set()
        self.calls: set[str] = set()
        self.names: set[str] = set()
        self.assigned_names: set[str] = set()
        self.assignments: dict[str, ast.AST] = {}

        for node in ast.walk(self.tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                self.functions[node.name] = node
            elif isinstance(node, ast.Import):
                self.imports.update(alias.name for alias in node.names)
            elif isinstance(node, ast.ImportFrom):
                if node.module:
                    self.imports.add(node.module)
            elif isinstance(node, ast.Call):
                name = dotted_name(node.func)
                if name:
                    self.calls.add(name)
            elif isinstance(node, ast.Name):
                self.names.add(node.id)
            elif isinstance(node, ast.Assign):
                for target in node.targets:
                    for name in target_names(target):
                        self.assigned_names.add(name)
                        self.assignments[name] = node.value
            elif isinstance(node, ast.AnnAssign):
                for name in target_names(node.target):
                    self.assigned_names.add(name)
                    if node.value is not None:
                        self.assignments[name] = node.value

        self.strings = strings_in(self.tree)
        self.string_text = "\n".join(self.strings)

    def function(self, name: str) -> ast.FunctionDef | ast.AsyncFunctionDef:
        return self.functions[name]

    def function_calls(self, name: str) -> set[str]:
        return calls_in(self.function(name))

    def function_names(self, name: str) -> set[str]:
        return names_in(self.function(name))

    def function_strings(self, name: str) -> list[str]:
        return strings_in(self.function(name))

    def assigned_integer(self, name: str) -> int | None:
        value = self.assignments.get(name)
        if isinstance(value, ast.Constant) and isinstance(value.value, int):
            return value.value
        return None


class MineSafeNonDestructiveRegressionTests(unittest.TestCase):
    """현재 MineSafe AI의 핵심 정적 계약을 확인한다."""

    TOTAL_TESTS = 14
    pass_count = 0
    warning_count = 0

    @classmethod
    def setUpClass(cls) -> None:
        cls.initial_exists = APP_PATH.is_file()
        cls.initial_size = APP_PATH.stat().st_size if cls.initial_exists else 0
        cls.initial_mtime_ns = APP_PATH.stat().st_mtime_ns if cls.initial_exists else 0
        cls.initial_hash = ""
        cls.model: StaticAppModel | None = None
        cls.load_error = ""

        if cls.initial_exists:
            raw = APP_PATH.read_bytes()
            cls.initial_hash = hashlib.sha256(raw).hexdigest().upper()
            try:
                cls.model = StaticAppModel(APP_PATH)
            except (OSError, UnicodeError, SyntaxError) as exc:
                cls.load_error = f"{type(exc).__name__}: {exc}"

    @classmethod
    def tearDownClass(cls) -> None:
        failed_or_errored = cls.TOTAL_TESTS - cls.pass_count
        print(
            f"[SUMMARY] 통과 {cls.pass_count} / 실패·오류 {failed_or_errored} "
            f"/ 경고 {cls.warning_count}",
            flush=True,
        )

    def require_model(self) -> StaticAppModel:
        self.assertTrue(self.initial_exists, "app.py 파일이 없습니다.")
        self.assertIsNotNone(
            self.model,
            f"app.py 정적 분석 모델을 만들 수 없습니다: {self.load_error}",
        )
        return self.model  # type: ignore[return-value]

    def pass_marker(self, message: str) -> None:
        type(self).pass_count += 1
        print(f"[PASS] {message}", flush=True)

    def warn_marker(self, message: str) -> None:
        type(self).warning_count += 1
        print(f"[WARN] {message}", flush=True)

    def assert_any_text_group(
        self,
        text: str,
        alternatives: tuple[str, ...],
        description: str,
    ) -> None:
        self.assertTrue(
            any(value in text for value in alternatives),
            f"{description} 구조를 찾지 못했습니다.",
        )

    def test_01_app_file_and_ast_syntax(self) -> None:
        self.assertTrue(self.initial_exists, "app.py 파일이 없습니다.")
        self.assertGreater(self.initial_size, 0, "app.py 파일이 비어 있습니다.")
        model = self.require_model()
        self.assertTrue(model.functions, "함수 정의를 찾지 못했습니다.")
        self.assertTrue(
            any(call.startswith("st.") for call in model.calls),
            "Streamlit 실행 코드 구조를 찾지 못했습니다.",
        )
        self.pass_marker("app.py 파일 존재, 비어 있지 않음, AST 문법 분석")

    def test_02_streamlit_webapp_structure(self) -> None:
        model = self.require_model()
        self.assertIn("streamlit", model.imports, "Streamlit import가 없습니다.")
        self.assertIn("st.set_page_config", model.calls)
        self.assertTrue(
            {"st.text_area", "st.text_input", "st.chat_input"} & model.calls,
            "사용자 질문 입력 구조가 없습니다.",
        )
        self.assertTrue(
            {"st.markdown", "st.write", "st.chat_message"} & model.calls,
            "답변 표시 구조가 없습니다.",
        )
        self.assertTrue(
            {"st.sidebar.radio", "st.sidebar.selectbox"} & model.calls,
            "사이드바 모드 선택 구조가 없습니다.",
        )
        self.assertIn("render_rag_result", model.functions)
        render_calls = model.function_calls("render_rag_result")
        self.assertTrue(
            {"st.markdown", "st.write"} & render_calls,
            "RAG 답변 렌더링 함수에 화면 출력 호출이 없습니다.",
        )
        self.pass_marker("Streamlit 질문 입력, 답변 표시, 사이드바/모드 선택 구조")

    def test_03_official_document_rag_structure(self) -> None:
        model = self.require_model()
        self.assertIn("chromadb", model.imports)
        self.assertIn("mine_safety_docs", model.string_text)
        self.assertIn("chromadb.PersistentClient", model.calls)
        self.assertTrue(
            any(call.endswith(".get_collection") for call in model.calls),
            "Chroma collection 조회 구조가 없습니다.",
        )
        self.assertTrue(
            any(call.endswith(".query") for call in model.calls),
            "Vector 검색 query 구조가 없습니다.",
        )
        required_functions = {
            "load_chroma_collection",
            "search_vector_db",
            "make_search_result",
            "build_context",
            "get_source",
            "get_chunk_id",
            "render_evidence_table",
        }
        self.assertTrue(required_functions.issubset(model.functions))
        context_text = "\n".join(model.function_strings("build_context"))
        self.assertIn("source", context_text)
        self.assertIn("chunk_id", context_text)
        rag_calls = model.function_calls("run_rag_flow")
        self.assertIn("search_vector_db", rag_calls)
        self.assertIn("generate_local_fallback_answer", rag_calls)
        self.pass_marker("공식 RAG collection, 검색 결과, source/chunk_id 및 답변 연결 구조")

    def test_04_answer_modes_and_fallback(self) -> None:
        model = self.require_model()
        mode_names = {"STABLE_MODE", "GEMINI_MODE", "HYBRID_MODE"}
        self.assertTrue(mode_names.issubset(model.assigned_names))
        run_node = model.function("run_rag_flow")
        run_calls = calls_in(run_node)
        run_names = names_in(run_node)
        self.assertIn("STABLE_MODE", run_names)
        self.assertIn("HYBRID_MODE", run_names)
        self.assertIn("generate_local_fallback_answer", run_calls)
        self.assertIn("generate_gemini_answer", run_calls)
        self.assertIn("generate_gemini_answer", model.functions)
        gemini_calls = model.function_calls("generate_gemini_answer")
        self.assertIn("generate_local_fallback_answer", gemini_calls)
        self.assertTrue(
            any(isinstance(node, ast.Try) for node in ast.walk(model.function("execute_gemini_request"))),
            "Gemini 오류 예외 처리 구조가 없습니다.",
        )
        status_text = "\n".join(model.function_strings("run_rag_flow")).lower()
        self.assertIn("fallback", status_text)
        self.pass_marker("안정형, Gemini 설명형, 하이브리드 및 실패 시 fallback 구조")

    def test_05_question_scope_classification(self) -> None:
        model = self.require_model()
        self.assertIn("detect_question_intent", model.functions)
        self.assertIn("build_out_of_scope_answer", model.functions)
        self.assertIn("OUT_OF_SCOPE_INTENT", model.assigned_names)
        intent_node = model.function("detect_question_intent")
        self.assertIn("OUT_OF_SCOPE_INTENT", names_in(intent_node))

        run_node = model.function("run_rag_flow")
        out_scope_branches = []
        for node in ast.walk(run_node):
            if isinstance(node, ast.If) and "OUT_OF_SCOPE_INTENT" in names_in(node.test):
                out_scope_branches.append(node)
        self.assertTrue(out_scope_branches, "범위 밖 질문 분기를 찾지 못했습니다.")
        self.assertTrue(
            any(has_return_in(branch) for branch in out_scope_branches),
            "범위 밖 질문의 별도 반환 구조가 없습니다.",
        )
        self.pass_marker("광산 안전 질문 분류와 범위 밖 질문 별도 안내 분기")

    def test_06_kras_risk_assessment_structure(self) -> None:
        model = self.require_model()
        required_functions = {
            "build_kras_intent_answer",
            "build_kras_risk_assessment_section",
            "infer_risk_level_for_kras",
            "build_risk_assessment_draft",
        }
        self.assertTrue(required_functions.issubset(model.functions))
        kras_text = "\n".join(
            model.function_strings("build_kras_risk_assessment_section")
            + model.function_strings("build_kras_intent_answer")
            + model.function_strings("build_risk_assessment_draft")
        )
        self.assert_any_text_group(kras_text, ("KRAS", "위험성평가"), "KRAS/위험성평가")
        self.assert_any_text_group(kras_text, ("위험요인", "유해·위험요인"), "위험요인")
        self.assert_any_text_group(kras_text, ("현재 위험성", "현재위험성", "현재 위험도"), "현재 위험성")
        self.assert_any_text_group(kras_text, ("감소대책", "감소 대책", "위험 감소"), "감소대책")
        self.assert_any_text_group(
            kras_text,
            ("조치 후 잔여위험성", "조치 후 위험성", "조치 후 위험도", "개선 후 위험"),
            "조치 후 위험도",
        )
        self.pass_marker("KRAS 위험요인, 현재 위험성, 감소대책, 조치 후 잔여위험성 구조")

    def test_07_evidence_record_recommendation(self) -> None:
        model = self.require_model()
        self.assertIn("recommend_evidence_records", model.functions)
        recommendation = model.function("recommend_evidence_records")
        recommendation_text = "\n".join(strings_in(recommendation))
        evidence_categories = (
            "작업중지",
            "점검표",
            "교육",
            "보호구",
            "사진",
            "측정",
        )
        category_count = sum(word in recommendation_text for word in evidence_categories)
        self.assertGreaterEqual(category_count, 3, "증빙자료 종류가 충분히 연결되어 있지 않습니다.")
        self.assertTrue(
            any(isinstance(node, ast.If) for node in ast.walk(recommendation)),
            "질문 내용과 증빙자료를 연결하는 조건 구조가 없습니다.",
        )
        self.assertIn("recommend_evidence_records", model.function_calls("render_rag_result"))
        self.assertIn("render_recommended_evidence_records", model.function_calls("render_rag_result"))
        self.pass_marker("질문 유형별 작업중지·점검·교육·보호구 등 증빙자료 추천 구조")

    def test_08_reference_case_warning_points(self) -> None:
        model = self.require_model()
        required_functions = {
            "match_reference_cases",
            "case_warning_points_for_intent",
            "format_reference_cases_for_prompt",
            "render_latest_reference_cases",
        }
        self.assertTrue(required_functions.issubset(model.functions))
        self.assertIn(
            "reference_case_categories_for_intent",
            model.function_calls("match_reference_cases"),
        )
        reference_text = "\n".join(
            model.function_strings("format_reference_cases_for_prompt")
            + model.function_strings("render_latest_reference_cases")
            + model.function_strings("case_warning_points_for_intent")
        )
        self.assert_any_text_group(
            reference_text,
            ("공식 법령 판단 근거가 아니라", "공식 법령 판단이 아니라", "참고 예시"),
            "공식 근거와 사례 참고자료 구분",
        )

        run_node = model.function("run_rag_flow")
        out_scope_if_lines = [
            node.lineno
            for node in ast.walk(run_node)
            if isinstance(node, ast.If) and "OUT_OF_SCOPE_INTENT" in names_in(node.test)
        ]
        match_case_lines = [
            node.lineno
            for node in ast.walk(run_node)
            if isinstance(node, ast.Call) and dotted_name(node.func) == "match_reference_cases"
        ]
        self.assertTrue(out_scope_if_lines and match_case_lines)
        self.assertLess(
            min(out_scope_if_lines),
            min(match_case_lines),
            "범위 밖 질문 조기 반환보다 사례 매칭이 먼저 실행될 수 있습니다.",
        )
        self.pass_marker("사례 주의 포인트, 공식 근거 구분 및 범위 밖 질문 조기 분기")

    def test_09_naver_news_static_structure(self) -> None:
        model = self.require_model()
        env_names = {
            "NAVER_CLIENT_ID",
            "NAVER_CLIENT_SECRET",
            "ENABLE_LIVE_CASE_SEARCH",
            "LIVE_CASE_SEARCH_PROVIDER",
        }
        for env_name in env_names:
            self.assertIn(env_name, model.string_text, f"{env_name} 이름이 없습니다.")

        self.assertEqual(model.assigned_integer("LIVE_CASE_SEARCH_TTL_SECONDS"), 1800)
        required_functions = {
            "normalize_naver_news_item",
            "search_naver_news_cases",
            "get_live_reference_cases",
            "render_live_news_reference_cases",
        }
        self.assertTrue(required_functions.issubset(model.functions))

        search_node = model.function("search_naver_news_cases")
        search_calls = calls_in(search_node)
        self.assertTrue(
            {"requests.get", "request.urlopen"} & search_calls,
            "네이버 뉴스 요청 처리 구조가 없습니다.",
        )
        exception_handlers = [
            handler
            for node in ast.walk(search_node)
            if isinstance(node, ast.Try)
            for handler in node.handlers
        ]
        self.assertTrue(
            any(contains_empty_list_return(handler) for handler in exception_handlers),
            "API 오류 시 빈 목록 fallback 구조가 없습니다.",
        )

        normalized_text = "\n".join(model.function_strings("normalize_naver_news_item"))
        for field_group in (("title",), ("description", "summary"), ("pubDate", "published"), ("link",)):
            self.assertTrue(
                any(field in normalized_text for field in field_group),
                f"뉴스 필드 처리 구조가 없습니다: {field_group[0]}",
            )

        cache_function = model.function("search_naver_news_cases")
        cache_decorators = [
            decorator
            for decorator in cache_function.decorator_list
            if isinstance(decorator, ast.Call) and dotted_name(decorator.func) == "st.cache_data"
        ]
        self.assertTrue(cache_decorators, "실시간 사례 검색 cache_data 구조가 없습니다.")
        ttl_names = set().union(*(names_in(item) for item in cache_decorators))
        self.assertIn("LIVE_CASE_SEARCH_TTL_SECONDS", ttl_names)

        live_text = "\n".join(
            model.function_strings("format_live_news_cases_for_prompt")
            + model.function_strings("render_live_news_reference_cases")
        )
        self.assert_any_text_group(
            live_text,
            ("공식 법령 판단 근거가 아닙니다", "공식 법령 판단 근거가 아니라"),
            "뉴스 참고자료 제한 안내",
        )
        self.pass_marker("네이버 뉴스 환경변수 이름, 요청/오류 fallback, 필드, 캐시와 1800초 TTL")

    def test_10_conservative_major_accident_law_wording(self) -> None:
        model = self.require_model()
        law_text = model.string_text
        conservative_groups = (
            ("쟁점이 될 수",),
            ("관리상 미흡으로 지적될 수",),
            ("안전보건관리체계 이행 여부", "안전보건관리체계가 실제로 작동"),
            ("실제 위반 여부는 사고 경위", "실제 위반 여부는"),
        )
        matched = sum(any(text in law_text for text in group) for group in conservative_groups)
        self.assertGreaterEqual(matched, 3, "보수적인 법령 표현 안전장치가 부족합니다.")
        self.assertIn("실제 위반 여부는", law_text)

        ambiguous = []
        for value in model.strings:
            if not any(term in value for term in ("위반 확정", "처벌 확정")):
                continue
            if not any(guard in value for guard in ("않", "말", "금지", "단정", "만들지")):
                ambiguous.append(value)
        if ambiguous:
            self.warn_marker("법적 확정 표현으로 오해될 수 있는 템플릿은 수동 문맥 확인 필요")
        self.pass_marker("중대재해처벌법 관련 보수적 표현과 실제 사실관계 확인 원칙")

    def test_11_no_direct_secret_output_structure(self) -> None:
        model = self.require_model()
        sensitive_pattern = re.compile(
            r"(?:api_?key|client_?id|client_?secret|password|access_?token|"
            r"secret_value|env_key|naver_client)",
            re.IGNORECASE,
        )
        sensitive_env_names = {
            "NAVER_CLIENT_ID",
            "NAVER_CLIENT_SECRET",
            "GEMINI_API_KEY",
        }

        def expression_reads_secret_source(node: ast.AST, tainted: set[str]) -> bool:
            if any(
                isinstance(item, ast.Name)
                and (item.id in tainted or sensitive_pattern.search(item.id))
                for item in ast.walk(node)
            ):
                return True
            for item in ast.walk(node):
                if isinstance(item, ast.Attribute) and dotted_name(item).startswith("st.secrets"):
                    return True
                if not isinstance(item, ast.Call):
                    continue
                call_name = dotted_name(item.func)
                if call_name in {"get_gemini_api_key", "os.getenv", "os.environ.get"}:
                    if call_name == "get_gemini_api_key":
                        return True
                    if item.args and isinstance(item.args[0], ast.Constant):
                        if item.args[0].value in sensitive_env_names:
                            return True
            return False

        dangerous_exact = {
            "print",
            "builtins.print",
            "st.write",
            "st.code",
            "st.text",
            "st.markdown",
            "st.caption",
            "st.json",
            "st.error",
            "st.warning",
            "st.info",
            "st.success",
            "st.sidebar.write",
            "st.sidebar.code",
            "st.sidebar.text",
            "st.sidebar.markdown",
            "st.sidebar.caption",
            "st.sidebar.json",
            "st.sidebar.error",
            "st.sidebar.warning",
            "st.sidebar.info",
            "st.sidebar.success",
        }

        def walk_one_scope(root: ast.AST):
            """중첩 함수 경계를 넘지 않고 한 실행 범위의 노드만 순회한다."""
            stack = list(ast.iter_child_nodes(root))
            while stack:
                node = stack.pop()
                if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda)):
                    continue
                yield node
                stack.extend(ast.iter_child_nodes(node))

        violations: list[tuple[int, str]] = []
        scopes: list[ast.AST] = [model.tree, *model.functions.values()]
        for scope in scopes:
            scope_nodes = list(walk_one_scope(scope))
            tainted: set[str] = set()
            assignments: list[tuple[set[str], ast.AST]] = []
            for node in scope_nodes:
                if isinstance(node, ast.Assign):
                    targets = set().union(*(target_names(target) for target in node.targets))
                    assignments.append((targets, node.value))
                elif isinstance(node, ast.AnnAssign) and node.value is not None:
                    assignments.append((target_names(node.target), node.value))

            changed = True
            while changed:
                changed = False
                for targets, value in assignments:
                    if expression_reads_secret_source(value, tainted):
                        new_names = targets - tainted
                        if new_names:
                            tainted.update(new_names)
                            changed = True

            for node in scope_nodes:
                if not isinstance(node, ast.Call):
                    continue
                call_name = dotted_name(node.func)
                is_logger = call_name.startswith(("logger.", "logging.")) and call_name.endswith(
                    (".debug", ".info", ".warning", ".error", ".exception")
                )
                if call_name not in dangerous_exact and not is_logger:
                    continue
                visible_values = list(node.args) + [keyword.value for keyword in node.keywords]
                if any(expression_reads_secret_source(value, tainted) for value in visible_values):
                    violations.append((getattr(node, "lineno", 0), call_name))

        safe_locations = ", ".join(f"line {line} {name}" for line, name in violations)
        self.assertFalse(
            violations,
            "인증정보 변수를 화면/로그에 직접 출력할 수 있는 구조가 있습니다: " + safe_locations,
        )
        self.pass_marker("환경변수 이름은 허용하되 인증정보 값을 직접 화면/로그에 출력하지 않는 구조")

    def test_12_external_reference_and_official_evidence_separation(self) -> None:
        model = self.require_model()
        render_calls = model.function_calls("render_rag_result")
        self.assertTrue(
            {"render_evidence_table", "render_evidence_card"} & render_calls,
            "공식 RAG 근거 렌더링 구조가 없습니다.",
        )
        self.assertIn("render_live_news_reference_cases", render_calls)
        self.assertIn("render_latest_reference_cases", render_calls)

        prompt_calls = model.function_calls("build_prompt")
        self.assertIn("build_context", prompt_calls)
        self.assertIn("format_reference_cases_for_prompt", prompt_calls)
        self.assertIn("format_live_news_cases_for_prompt", prompt_calls)

        run_strings = set(model.function_strings("run_rag_flow"))
        self.assertIn("reference_cases", run_strings)
        self.assertIn("live_news_cases", run_strings)
        self.assertIn("question_intent", run_strings)
        self.pass_marker("공식 RAG 근거, 사례 주의 포인트, 뉴스 참고자료의 분리된 데이터/렌더링 구조")

    def test_13_reference_ui_diagnostics(self) -> None:
        model = self.require_model()
        reference_checks = {
            "MineSafe AI 제목": "MineSafe AI" in model.string_text,
            "현장/관리자 모드 명칭": (
                "현장관리자 모드" in model.string_text
                and "관리자/개발자 모드" in model.string_text
            ),
            "직접 질문/시나리오 탭": (
                "직접 질문" in model.string_text
                and "질문 시나리오 테스트 모드" in model.string_text
            ),
            "근거 상세 expander": "근거 문서 상세 보기" in model.string_text,
            "대시보드 CSS 계열": any("portal-" in value for value in model.strings),
        }
        for description, present in reference_checks.items():
            if not present:
                self.warn_marker(f"{description} 표현 또는 배치는 변경되어 수동 확인 필요")
        self.pass_marker("UI 문구·배치·CSS 참고 검사 완료(누락은 경고만 표시)")

    def test_99_app_integrity_unchanged(self) -> None:
        self.assertTrue(APP_PATH.is_file(), "테스트 중 app.py가 사라졌습니다.")
        current_raw = APP_PATH.read_bytes()
        current_hash = hashlib.sha256(current_raw).hexdigest().upper()
        current_stat = APP_PATH.stat()
        self.assertEqual(len(current_raw), self.initial_size, "app.py 크기가 바뀌었습니다.")
        self.assertEqual(current_stat.st_mtime_ns, self.initial_mtime_ns, "app.py 수정 시각이 바뀌었습니다.")
        self.assertEqual(current_hash, self.initial_hash, "app.py SHA-256이 바뀌었습니다.")
        self.pass_marker("app.py 무결성 유지(크기, 수정 시각, SHA-256 동일)")


if __name__ == "__main__":
    unittest.main(verbosity=2)
