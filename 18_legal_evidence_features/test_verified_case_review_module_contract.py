from __future__ import annotations

import ast
import importlib
import importlib.util
import json
from pathlib import Path
import re
import subprocess
import sys
from typing import Any
import unittest


ROOT_DIR = Path(__file__).resolve().parents[1]
APP_PATH = ROOT_DIR / "app.py"
REVIEW_PATH = ROOT_DIR / "verified_case_review.py"


def load_review_module():
    spec = importlib.util.spec_from_file_location(
        "verified_case_review_contract_target", REVIEW_PATH
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("verified_case_review 모듈 사양을 만들 수 없습니다.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class VerifiedCaseReviewModuleContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app_source = APP_PATH.read_text(encoding="utf-8-sig")
        cls.review_source = REVIEW_PATH.read_text(encoding="utf-8-sig")
        cls.app_tree = ast.parse(cls.app_source, filename=str(APP_PATH))
        cls.review_tree = ast.parse(cls.review_source, filename=str(REVIEW_PATH))
        cls.review = load_review_module()
        cls.app_attributes = {
            node.attr
            for node in ast.walk(cls.app_tree)
            if isinstance(node, ast.Attribute)
            and isinstance(node.value, ast.Name)
            and node.value.id == "verified_review"
        }

    def test_01_fresh_import_resolves_to_project_module(self) -> None:
        script = (
            "import json, pathlib, verified_case_review as v; "
            "print(json.dumps({'module_file': str(pathlib.Path(v.__file__).resolve()), "
            "'tier': v.AUTO_SCREENED_PUBLIC_TIER}, ensure_ascii=False))"
        )
        completed = subprocess.run(
            [sys.executable, "-B", "-c", script],
            cwd=ROOT_DIR,
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        payload = json.loads(completed.stdout.strip())
        self.assertEqual(REVIEW_PATH.resolve(), Path(payload["module_file"]).resolve())
        self.assertEqual("auto_screened", payload["tier"])

    def test_02_app_uses_expected_module_alias(self) -> None:
        imports = [
            alias
            for node in self.app_tree.body
            if isinstance(node, ast.Import)
            for alias in node.names
            if alias.name == "verified_case_review"
        ]
        self.assertEqual(1, len(imports))
        self.assertEqual("verified_review", imports[0].asname)

    def test_03_every_app_reference_exists_at_runtime(self) -> None:
        missing = sorted(
            name for name in self.app_attributes if not hasattr(self.review, name)
        )
        self.assertEqual([], missing)

    def test_04_required_contract_covers_every_app_reference(self) -> None:
        assignment = next(
            node
            for node in self.app_tree.body
            if isinstance(node, ast.Assign)
            and any(
                isinstance(target, ast.Name)
                and target.id == "VERIFIED_REVIEW_REQUIRED_ATTRIBUTES"
                for target in node.targets
            )
        )
        required = set(ast.literal_eval(assignment.value))
        self.assertEqual(self.app_attributes, required)

    def test_05_auto_screened_public_tier_is_canonical_alias(self) -> None:
        self.assertEqual("auto_screened", self.review.AUTO_SCREENED_STATUS)
        self.assertEqual(
            self.review.AUTO_SCREENED_STATUS,
            self.review.AUTO_SCREENED_PUBLIC_TIER,
        )

    def test_06_all_public_tier_constants_have_allowed_values(self) -> None:
        expected = {"verified", "auto_screened", "text_safe_fallback", "hidden"}
        actual = {
            self.review.VERIFIED_PUBLIC_TIER,
            self.review.AUTO_SCREENED_PUBLIC_TIER,
            self.review.TEXT_SAFE_FALLBACK_TIER,
            self.review.HIDDEN_PUBLIC_TIER,
        }
        self.assertEqual(expected, actual)
        self.assertEqual(expected, set(self.review.PUBLIC_CASE_TIERS))

    def test_07_status_and_public_tier_domains_are_consistent(self) -> None:
        self.assertIn(self.review.VERIFIED_PUBLIC_TIER, self.review.CASE_STATUSES)
        self.assertIn(self.review.AUTO_SCREENED_PUBLIC_TIER, self.review.CASE_STATUSES)
        self.assertNotIn(self.review.TEXT_SAFE_FALLBACK_TIER, self.review.CASE_STATUSES)
        self.assertNotIn(self.review.HIDDEN_PUBLIC_TIER, self.review.CASE_STATUSES)

    def test_08_hot_reload_contract_guard_exists(self) -> None:
        required_assignment = next(
            node
            for node in self.app_tree.body
            if isinstance(node, ast.Assign)
            and any(
                isinstance(target, ast.Name)
                and target.id == "VERIFIED_REVIEW_REQUIRED_ATTRIBUTES"
                for target in node.targets
            )
        )
        guard = next(
            node
            for node in self.app_tree.body
            if isinstance(node, ast.FunctionDef)
            and node.name == "ensure_verified_review_contract"
        )
        called_names = {
            node.func.attr
            for node in ast.walk(guard)
            if isinstance(node, ast.Call)
            and isinstance(node.func, ast.Attribute)
            and isinstance(node.func.value, ast.Name)
            and node.func.value.id == "importlib"
        }
        self.assertEqual({"invalidate_caches", "reload"}, called_names)
        self.assertIn(
            "__file__",
            {node.id for node in ast.walk(guard) if isinstance(node, ast.Name)},
        )
        namespace = {
            "__file__": str(APP_PATH),
            "Any": Any,
            "Path": Path,
            "importlib": importlib,
        }
        contract_module = ast.Module(
            body=[required_assignment, guard], type_ignores=[]
        )
        exec(compile(contract_module, str(APP_PATH), "exec"), namespace)
        runtime_review = importlib.import_module("verified_case_review")
        delattr(runtime_review, "AUTO_SCREENED_PUBLIC_TIER")
        refreshed = namespace["ensure_verified_review_contract"](runtime_review)
        self.assertIs(runtime_review, refreshed)
        self.assertEqual("auto_screened", refreshed.AUTO_SCREENED_PUBLIC_TIER)

    def test_09_import_has_no_top_level_api_or_vector_db_open(self) -> None:
        forbidden_calls = {"PersistentClient", "HttpClient", "urlopen", "request", "get", "post"}
        found: set[str] = set()
        for statement in self.review_tree.body:
            if isinstance(statement, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                continue
            for node in ast.walk(statement):
                if not isinstance(node, ast.Call):
                    continue
                if isinstance(node.func, ast.Name):
                    name = node.func.id
                elif isinstance(node.func, ast.Attribute):
                    name = node.func.attr
                else:
                    continue
                if name in forbidden_calls:
                    found.add(name)
        self.assertEqual(set(), found)

    def test_10_contract_code_does_not_render_secret_values(self) -> None:
        combined = f"{self.review_source}\n{ast.unparse(next(node for node in self.app_tree.body if isinstance(node, ast.FunctionDef) and node.name == 'ensure_verified_review_contract'))}"
        self.assertIsNone(
            re.search(
                r"(?i)(print|write|error|warning|info|markdown)\s*\([^\n]*(api[_ ]?key|client[_ ]?secret)",
                combined,
            )
        )


if __name__ == "__main__":
    unittest.main()
