from __future__ import annotations

import ast
import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch


ROOT_DIR = Path(__file__).resolve().parents[1]
APP_PATH = ROOT_DIR / "app.py"
REVIEW_PATH = ROOT_DIR / "verified_case_review.py"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"모듈을 읽을 수 없습니다: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def function_source(source: str, name: str) -> str:
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return ast.get_source_segment(source, node) or ""
    return ""


class PresentationCaseRecoveryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.review = load_module("presentation_case_recovery_test", REVIEW_PATH)
        cls.app_source = APP_PATH.read_text(encoding="utf-8-sig")

    def make_records(self):
        groups = (
            ("CONV", "컨베이어 끼임 사고"),
            ("VEH", "차량 후진 충돌 사고"),
            ("ROOF", "낙반 붕괴 매몰 사고"),
            ("ELEC", "전기 감전 사고"),
        )
        records = []
        for prefix, text in groups:
            for index in range(3):
                case_id = f"{prefix}-{index}"
                records.append(
                    {
                        "case_id": case_id,
                        "content_hash": f"hash-{case_id}",
                        "verification_status": "unverified",
                        "official_case": True,
                        "mine_relevance": "medium",
                        "ocr_quality_status": "pass",
                        "needs_manual_review": False,
                        "needs_reocr": False,
                        "text_quality_score": 100,
                        "reading_order_score": 100,
                        "metadata_quality_score": 65,
                        "source_document": "공식 자료",
                        "source_file": "source.pdf",
                        "source_period": "2025",
                        "page_start": 1,
                        "page_end": 1,
                        "industry": "",
                        "accident_type": "",
                        "accident_summary": text,
                        "display_accident_summary": text,
                        "display_cause_summary": "",
                        "display_prevention_summary": "",
                    }
                )
        records.append(
            {
                "case_id": "OUTSIDE",
                "content_hash": "hash-outside",
                "verification_status": "unverified",
                "accident_summary": "발표 유형과 무관한 사례",
            }
        )
        return records

    def test_01_priority_selection_is_limited_to_twelve(self):
        selected = self.review.priority_review_candidates(self.make_records())
        self.assertEqual(12, len(selected))
        self.assertNotIn("OUTSIDE", {item["case_id"] for item in selected})

    def test_02_recovery_connects_an_exact_source_card_image(self):
        from PIL import Image

        records = self.make_records()
        with tempfile.TemporaryDirectory() as temp_name:
            temp_root = Path(temp_name)
            docs = temp_root / "docs"
            docs.mkdir()
            pdf_path = docs / "source.pdf"
            pdf_path.write_bytes(b"official source fixture")
            expected_hash = hashlib.sha256(pdf_path.read_bytes()).hexdigest()
            output_dir = temp_root / "cards"
            audit_path = temp_root / "audit.jsonl"

            def fake_render(_pdf, _page, output_path, _renderer):
                Image.new("RGB", (2000, 1200), "white").save(output_path)

            manifest = {
                "source.pdf": {
                    "saved_filename": "source.pdf",
                    "source_id": "MOEL-SIREN-2025-BOOK",
                    "title": "공식 자료",
                    "publisher": "고용노동부",
                    "source_page_url": "https://www.moel.go.kr/official",
                    "coverage_start": "2025-01-01",
                    "coverage_end": "2025-12-31",
                    "page_count": 1,
                    "sha256": expected_hash,
                }
            }
            specs = [
                {
                    "case_id": "CONV-0",
                    "source_file": "source.pdf",
                    "page_start": 1,
                    "crop_box": [0.0, 0.0, 0.5, 1.0],
                    "industry": "제조업",
                    "industry_source": "original_card_image",
                    "accident_date": "2025-01-13",
                    "accident_type": "끼임",
                    "display_accident_summary": (
                        "2025년 1월 13일 제조업 사업장에서 근로자가 컨베이어 설비를 점검하던 중 "
                        "회전체에 끼어 사망했습니다."
                    ),
                }
            ]
            with (
                patch.object(self.review, "ROOT_DIR", temp_root),
                patch.object(self.review, "SOURCE_PDF_DIR", docs),
                patch.object(self.review, "PRESENTATION_RECOVERY_AUDIT_PATH", audit_path),
                patch.object(self.review, "_render_page_image", side_effect=fake_render),
            ):
                recovered, summary = self.review.recover_presentation_priority_cases(
                    records,
                    specs,
                    manifest_index=manifest,
                    output_dir=output_dir,
                    pdftoppm_path=Path("fake-pdftoppm"),
                    persist_events=False,
                )
                recovered_by_id = {item["case_id"]: item for item in recovered}
                restored = recovered_by_id["CONV-0"]
                self.assertEqual(1, summary["recovered_case_count"])
                self.assertTrue((temp_root / restored["original_page_image"]).is_file())
                self.assertEqual("제조업", restored["industry"])
                self.assertEqual("끼임", restored["accident_type"])
                self.assertEqual(records[-1], recovered_by_id["OUTSIDE"])

    def test_03_manifest_industry_requires_exact_source_filename(self):
        entry = {"saved_filename": "construction.pdf", "source_id": "MOEL-SIREN-2026-01-CONSTRUCTION"}
        self.assertEqual("건설업", self.review.official_industry_from_manifest("construction.pdf", entry))
        self.assertEqual("", self.review.official_industry_from_manifest("other.pdf", entry))
        book = {"saved_filename": "book.pdf", "source_id": "MOEL-SIREN-2025-BOOK"}
        self.assertEqual("", self.review.official_industry_from_manifest("book.pdf", book))

    def test_04_accident_type_must_be_explicit_in_display_summary(self):
        summary = "작업자가 후진 중인 굴착기에 부딪혀 사망했습니다."
        self.assertEqual("부딪힘", self.review.explicit_accident_type_for_recovery(summary, "부딪힘"))
        self.assertEqual("", self.review.explicit_accident_type_for_recovery(summary, "감전"))

    def test_05_corrupted_text_remains_excluded(self):
        self.assertTrue(self.review.detect_corrupted_ocr_text("안전 Zool XSF SCHRHON PuwBct 점검"))
        source = function_source(REVIEW_PATH.read_text(encoding="utf-8-sig"), "recover_presentation_priority_cases")
        self.assertIn("display_text_corruption_reasons", source)
        self.assertNotIn("replace_corrupted", source)

    def test_06_quality_thresholds_are_not_lowered(self):
        self.assertEqual(80, self.review.AUTO_SCREENED_MIN_TEXT_QUALITY)
        self.assertEqual(80, self.review.AUTO_SCREENED_MIN_READING_ORDER)
        self.assertEqual(70, self.review.AUTO_SCREENED_MIN_METADATA_QUALITY)
        self.assertEqual((40, 450), (self.review.AUTO_SCREENED_MIN_DISPLAY_CHARS, self.review.AUTO_SCREENED_MAX_DISPLAY_CHARS))

    def test_07_auto_screened_database_is_separate(self):
        names = {
            self.review.LAW_COLLECTION_NAME,
            self.review.SOURCE_CASE_COLLECTION_NAME,
            self.review.VERIFIED_CASE_COLLECTION_NAME,
            self.review.AUTO_SCREENED_CASE_COLLECTION_NAME,
        }
        self.assertEqual(4, len(names))

    def test_08_verified_priority_is_preserved(self):
        self.assertLess(
            self.review.PUBLIC_RANK_ORDER[("direct", "verified")],
            self.review.PUBLIC_RANK_ORDER[("direct", "auto_screened")],
        )
        self.assertLess(
            self.review.PUBLIC_RANK_ORDER[("analogous", "verified")],
            self.review.PUBLIC_RANK_ORDER[("analogous", "auto_screened")],
        )
        vehicle_squeeze = {
            "accident_type": "끼임",
            "display_accident_summary": "화물차량이 후진하다 접안시설 사이에 끼인 사고입니다.",
        }
        relation, _ = self.review.classify_case_relation(
            vehicle_squeeze,
            ("컨베이어", "벨트", "끼임", "말림"),
            ("회전체", "정비 중 재가동", "에너지 차단"),
        )
        self.assertEqual("", relation)
        stationary_excavator = {
            "equipment": "굴착기",
            "display_accident_summary": "채석장에서 굴착기로 상차하던 중 매몰된 사고입니다.",
        }
        relation, _ = self.review.classify_case_relation(
            stationary_excavator,
            ("후진", "신호수", "덤프트럭", "굴착기", "지게차", "차량 충돌"),
            ("이동식 장비", "사각지대", "깔림", "중장비", "작업자 충돌"),
        )
        self.assertEqual("", relation)

    def test_09_tab_order_is_preserved(self):
        expected = '["최근 뉴스 참고", "공식 재해사례", "핵심 주의사항"]'
        self.assertIn(expected, self.app_source)

    def test_10_out_of_scope_returns_before_case_database_load(self):
        source = function_source(self.app_source, "search_official_siren_cases")
        guard = source.index("question_type == OUT_OF_SCOPE_INTENT")
        database_load = source.index("load_official_case_collection")
        self.assertLess(guard, database_load)

    def test_11_public_card_still_uses_display_fields_only(self):
        source = function_source(self.app_source, "render_official_siren_case_card")
        for field in ("display_accident_summary", "display_cause_summary", "display_prevention_summary"):
            self.assertIn(field, source)
        self.assertNotIn("layout_ocr_text", source)

    def test_12_plan_loader_rejects_more_than_twelve_or_duplicate_ids(self):
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "plan.json"
            path.write_text(json.dumps({"cases": [{"case_id": "A"}, {"case_id": "A"}]}), encoding="utf-8")
            with self.assertRaises(self.review.ReviewWorkflowBlocked):
                self.review.load_presentation_recovery_plan(path)
            path.write_text(
                json.dumps({"cases": [{"case_id": f"C{index}"} for index in range(13)]}),
                encoding="utf-8",
            )
            with self.assertRaises(self.review.ReviewWorkflowBlocked):
                self.review.load_presentation_recovery_plan(path)

    def test_13_persisted_priority_set_does_not_drift_after_metadata_changes(self):
        records = self.make_records()
        initial = self.review.priority_review_candidates(records)
        initial_groups = {
            item["case_id"]: item["priority_review_group"]
            for item in initial
        }
        for record in records:
            if record["case_id"] in initial_groups:
                record["priority_review_group"] = initial_groups[record["case_id"]]
        outside = next(record for record in records if record["case_id"] == "OUTSIDE")
        outside.update(
            {
                "accident_type": "끼임",
                "accident_summary": "컨베이어 회전체에 끼이고 말린 사고",
                "verification_status": "unverified",
            }
        )
        after_metadata_change = self.review.priority_review_candidates(records)
        self.assertEqual(
            set(initial_groups),
            {item["case_id"] for item in after_metadata_change},
        )


if __name__ == "__main__":
    unittest.main()
