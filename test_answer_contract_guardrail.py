from __future__ import annotations

import unittest

from answer_contract import (
    AnswerContract,
    build_answer_contract,
    build_complete_core_judgment,
    build_rule_based_fallback,
    guard_answer_generation,
    incomplete_response_reason,
    official_case_status_message,
    validate_answer,
)
from query_understanding import analyze_query


class AnswerContractGuardrailTests(unittest.TestCase):
    def test_01_order_request_has_safety_sequence(self) -> None:
        understanding = analyze_query("천반 출수가 늘 때 무엇부터 해야 하는지 순서대로 알려줘")
        contract = build_answer_contract(understanding)
        labels = [group[0] for group in contract.required_sequence]
        self.assertEqual(labels[:4], ["작업중지", "대피", "출입통제", "보고"])
        self.assertIn("작업 재개", labels[-2])
        self.assertIn("기록", labels[-1])

    def test_02_simple_education_does_not_force_eight_steps(self) -> None:
        understanding = analyze_query("분진이 왜 위험한지 TBM 교육자료로 설명해 줘")
        contract = build_answer_contract(understanding)
        self.assertEqual(contract.required_sequence, [])
        self.assertIn("위험 원인", contract.required_sections)
        self.assertIn("교육·TBM 공유사항", contract.required_sections)

    def test_03_restart_request_requires_restart_condition(self) -> None:
        understanding = analyze_query("환기 복구 뒤 다시 들어가려면 무엇을 확인해야 해?")
        contract = build_answer_contract(understanding)
        self.assertTrue(contract.restart_required)
        self.assertIn("작업 재개 조건", contract.required_sections)

    def test_04_case_request_has_case_status_contract(self) -> None:
        understanding = analyze_query("컨베이어 끼임 실제 사고사례도 알려줘")
        contract = build_answer_contract(understanding)
        self.assertEqual(contract.requested_case_mode, "official")
        self.assertIn("공식 사고사례", contract.required_sections)

    def test_05_no_direct_case_does_not_force_unrelated_case(self) -> None:
        understanding = analyze_query("컨베이어 끼임 실제 사고사례도 알려줘")
        message = official_case_status_message(understanding, [])
        self.assertIn("직접 관련 공식 사례는 확인되지 않았습니다", message)
        self.assertIn("무관한 사례는 표시하지 않습니다", message)

    def test_06_evidence_request_requires_real_source_and_chunk_id(self) -> None:
        understanding = analyze_query("전기 정비의 공식 법적 근거는 뭐야?")
        contract = build_answer_contract(understanding)
        result = {"source": "공식 전기안전 지침", "chunk_id": "elec-01", "text": "전원 차단과 점검"}
        missing = validate_answer(
            "## 관련 근거\n공식 근거 보완 필요 안내입니다. 전기설비 정비 전 전원 차단과 점검이 필요합니다.",
            contract,
            understanding,
            evidence_results=[result],
        )
        cited = validate_answer(
            "## 관련 근거 문서\n공식 전기안전 지침 (chunk_id: elec-01)\n전기설비 정비 전 전원 차단과 점검이 필요합니다.",
            contract,
            understanding,
            evidence_results=[result],
        )
        self.assertIn("실제 문서명·chunk_id 근거", missing.missing_sections)
        self.assertNotIn("실제 문서명·chunk_id 근거", cited.missing_sections)

    def test_07_unsupported_important_number_is_warned(self) -> None:
        understanding = analyze_query("가스가 의심될 때 어떻게 해야 해?")
        contract = build_answer_contract(understanding)
        validation = validate_answer(
            "작업을 중지하고 대피합니다. 근거 없이 가스 농도 5%가 될 때까지 기다린 뒤 책임자가 재개를 승인합니다.",
            contract,
            understanding,
        )
        self.assertIn("5%", validation.unsupported_number_warnings)

    def test_08_unrelated_dominant_topic_requests_supplement(self) -> None:
        understanding = analyze_query("강우 뒤 발파공 물을 장약 전에 어떻게 점검해?")
        contract = build_answer_contract(understanding)
        validation = validate_answer(
            "불발공 접근 금지와 불발공 감시가 핵심입니다. 불발 절차만 반복하며 불발 처리 후 기록합니다.",
            contract,
            understanding,
        )
        self.assertTrue(validation.retry_recommended)
        self.assertTrue(validation.unrelated_topic_warnings or validation.missing_concepts)

    def test_09_one_regeneration_then_fallback(self) -> None:
        understanding = analyze_query("환기 정지 때 무엇부터 해야 해?")
        contract = AnswerContract(allowed_retry_count=1)
        calls = []

        def regenerate(_validation):
            calls.append("retry")
            return "제공된 [안정"

        result = guard_answer_generation(
            "제공된 [안정",
            understanding,
            contract,
            regenerate=regenerate,
            fallback_factory=lambda: (
                "작업을 중지하고 작업자를 안전한 장소로 대피시킵니다. "
                "위험구역을 통제하고 책임자 확인 후 조치와 재개 판단을 기록합니다."
            ),
        )
        self.assertEqual(calls, ["retry"])
        self.assertEqual(result.generation_calls, 2)
        self.assertTrue(result.fallback_used)

    def test_10_incomplete_fragment_is_blocked(self) -> None:
        self.assertTrue(incomplete_response_reason("제공된 [안정"))
        self.assertTrue(incomplete_response_reason("작업자는 즉시,"))
        self.assertEqual(
            incomplete_response_reason(
                "- 작업을 중지하고 대피\n- 위험구역 출입통제\n- 책임자 확인 후 작업 재개 판단과 기록"
            ),
            "",
        )

    def test_11_retry_count_never_becomes_unbounded(self) -> None:
        understanding = analyze_query("전기설비 정비 전에 무엇을 해야 해?")
        calls = []
        result = guard_answer_generation(
            "",
            understanding,
            AnswerContract(allowed_retry_count=1),
            regenerate=lambda _validation: calls.append(1) or "",
            fallback_factory=lambda: "전원을 차단하고 잠금·표지 후 무전압을 확인합니다. 책임자 확인과 기록 뒤 재개를 판단합니다.",
        )
        self.assertEqual(len(calls), 1)
        self.assertEqual(result.generation_calls, 2)

    def test_12_stable_fallback_and_core_judgment_are_complete(self) -> None:
        understanding = analyze_query("천반 출수가 늘고 지보가 휘면 작업자부터 어디로 보내?")
        contract = build_answer_contract(understanding)
        answer = build_rule_based_fallback(understanding, contract)
        judgment = build_complete_core_judgment(understanding)
        self.assertTrue(answer.strip())
        self.assertEqual(incomplete_response_reason(answer), "")
        self.assertEqual(incomplete_response_reason(judgment), "")
        self.assertLessEqual(judgment.count("."), 2)


if __name__ == "__main__":
    unittest.main()
