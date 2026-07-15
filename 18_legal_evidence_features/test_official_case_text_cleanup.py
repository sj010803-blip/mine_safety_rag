import ast
import importlib.util
import re
import tempfile
import unittest
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PIPELINE_PATH = ROOT / "21_official_accident_case_pipeline" / "build_official_siren_case_db.py"
APP_PATH = ROOT / "app.py"
PIPELINE_SOURCE = PIPELINE_PATH.read_text(encoding="utf-8-sig")
APP_SOURCE = APP_PATH.read_text(encoding="utf-8-sig")
ast.parse(PIPELINE_SOURCE)
ast.parse(APP_SOURCE)

spec = importlib.util.spec_from_file_location("official_case_cleanup_pipeline", PIPELINE_PATH)
pipeline = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(pipeline)


def sample_case(**updates):
    case = {
        "case_id": "KOSHA-SIREN-TEST-0001",
        "content_hash": "a" * 64,
        "official_case": True,
        "source_document": "공식 중대재해사이렌",
        "source_file": "26년 중대재해사이렌(제조).pdf",
        "page_start": 3,
        "page_end": 3,
        "accident_date": "2026년 1월 2일",
        "industry": "제조업",
        "accident_type": "끼임",
        "accident_summary": "사고 개요: 근로자가 설비를 점검하던 중 회전체에 끼여 사고가 발생했습니다.",
        "cause_summary": "",
        "prevention_summary": "",
        "mine_relevance": "medium",
        "ocr_quality_status": "pass",
    }
    case.update(updates)
    return case


class OfficialCaseTextCleanupTests(unittest.TestCase):
    def test_01_ocr_normalization_functions_exist(self):
        for name in ("normalize_ocr_text", "remove_ocr_noise", "repair_korean_spacing"):
            self.assertTrue(callable(getattr(pipeline, name, None)))

    def test_02_known_english_ocr_noise_is_removed(self):
        cleaned = pipeline.remove_ocr_noise("안전 Zool XSF SCHRHON PuwBct 점검")
        self.assertNotRegex(cleaned, r"Zool|XSF|SCHRHON|PuwBct")

    def test_03_safe_english_abbreviations_are_preserved(self):
        cleaned = pipeline.remove_ocr_noise("TBM PPE CCTV LOTO 점검")
        for token in ("TBM", "PPE", "CCTV", "LOTO"):
            self.assertIn(token, cleaned)

    def test_04_multicolumn_reading_order_function_exists(self):
        self.assertTrue(callable(pipeline.order_ocr_blocks))

    def test_05_blocks_are_ordered_by_column_coordinates(self):
        blocks = [
            {"text": "오른쪽", "x0": 700, "y0": 10},
            {"text": "왼쪽 아래", "x0": 10, "y0": 200},
            {"text": "왼쪽 위", "x0": 10, "y0": 20},
        ]
        ordered = pipeline.order_ocr_blocks(blocks, 900)
        self.assertEqual(["왼쪽 위", "왼쪽 아래", "오른쪽"], [item["text"] for item in ordered])

    def test_06_mixed_accident_blocks_are_detected(self):
        text = "2025년 1월 1일 사고 개요 첫 사례\n2025년 2월 2일 사고 개요 둘째 사례"
        self.assertTrue(pipeline.detect_mixed_case_blocks(text))

    def test_07_accident_section_boundaries_are_split(self):
        result = pipeline.split_accident_sections(
            "사고 개요: 설비 점검 중 끼임 사고가 발생했습니다.\n발생 원인: 전원 미차단\n예방대책: 전원 차단"
        )
        self.assertIn("끼임", result["accident_summary"])
        self.assertNotIn("예방대책", result["accident_summary"])

    def test_08_one_character_industry_is_not_saved(self):
        self.assertEqual("", pipeline.normalize_industry("인", "관련 업종 표기 없음"))

    def test_09_accident_type_can_use_explicit_source_word(self):
        self.assertEqual("감전", pipeline.infer_accident_type_from_text("정보 없음", "전기 점검 중 감전 사고"))

    def test_10_full_and_display_summaries_are_separate(self):
        cleaned = pipeline.clean_case_record(sample_case())
        self.assertIn("full_accident_summary", cleaned)
        self.assertIn("display_accident_summary", cleaned)

    def test_11_display_accident_summary_has_length_limit(self):
        case = sample_case(accident_summary="사고 개요: " + ("근로자가 설비를 점검했습니다. " * 100))
        cleaned = pipeline.clean_case_record(case)
        self.assertLessEqual(len(cleaned["display_accident_summary"]), pipeline.DISPLAY_SUMMARY_MAX_CHARS + 1)

    def test_12_missing_cause_and_prevention_use_safe_display_messages(self):
        cleaned = pipeline.clean_case_record(sample_case())
        self.assertIn("별도 원인", cleaned["display_cause_summary"])
        self.assertIn("별도 예방사항", cleaned["display_prevention_summary"])

    def test_13_missing_source_sections_are_not_invented(self):
        cleaned = pipeline.clean_case_record(sample_case())
        self.assertEqual("", cleaned["cause_summary"])
        self.assertEqual("", cleaned["prevention_summary"])

    def test_14_low_quality_case_is_excluded_from_db(self):
        case = sample_case(
            accident_summary=(
                "근로자가 설비를 점검하던 중 위험구역에 접근하여 끼임 사고가 발생했습니다. "
                "동료 작업자가 설비를 정지하고 현장 관리자에게 즉시 사고 상황을 보고했습니다."
            ),
            text_quality_score=10,
            needs_manual_review=False,
        )
        reason = pipeline.case_vector_db_exclusion_reason(case, set(), set())
        self.assertEqual("low_text_quality", reason)

    def test_15_candidate_db_is_verified_before_swap_plan(self):
        plan = pipeline.validated_candidate_swap_plan(datetime(2030, 1, 2, 3, 4, 5))
        self.assertIn("candidate", plan)
        self.assertIn("backup", plan)
        self.assertIn("20300102_030405", plan["backup"])

    def test_16_legal_vector_db_path_remains_separate(self):
        self.assertNotEqual(pipeline.CASE_VECTOR_DB_DIR.resolve(), pipeline.LAW_VECTOR_DB_DIR.resolve())
        self.assertIn("기존 법령 DB는 교체 대상이 아닙니다", PIPELINE_SOURCE)

    def test_17_out_of_scope_guard_precedes_case_db_load(self):
        start = APP_SOURCE.index("def search_official_siren_cases(")
        end = APP_SOURCE.index("\ndef ", start + 10)
        body = APP_SOURCE[start:end]
        self.assertLess(body.index("OUT_OF_SCOPE_INTENT"), body.index("load_official_case_collection"))

    def test_18_official_case_ui_and_additional_expander_remain(self):
        self.assertIn("render_official_siren_cases", APP_SOURCE)
        self.assertIn("추가 사례", APP_SOURCE)

    def test_19_app_prefers_display_fields(self):
        self.assertRegex(APP_SOURCE, r"display_accident_summary.+or case\.get\(\"accident_summary\"")
        self.assertIn("원문 OCR 내용 보기", APP_SOURCE)

    def test_20_secret_values_are_not_rendered(self):
        dangerous = re.compile(r"st\.(?:write|text|markdown)\([^\n]*(?:os\.getenv|environ\[)")
        self.assertIsNone(dangerous.search(APP_SOURCE))


if __name__ == "__main__":
    unittest.main()
