from __future__ import annotations

from dataclasses import asdict, dataclass, field
import re
from typing import Any, Callable, Iterable

from query_understanding import (
    DOMAIN_BLASTING,
    DOMAIN_CONVEYOR,
    DOMAIN_DUST,
    DOMAIN_ELECTRICAL,
    DOMAIN_EMERGENCY,
    DOMAIN_GROUND,
    DOMAIN_MOBILE,
    DOMAIN_PPE,
    DOMAIN_REPORTING,
    DOMAIN_RISK,
    DOMAIN_TBM,
    DOMAIN_VENTILATION,
    QueryUnderstanding,
)


@dataclass(frozen=True)
class AnswerContract:
    required_sections: list[str] = field(default_factory=list)
    required_concept_groups: list[list[str]] = field(default_factory=list)
    required_sequence: list[list[str]] = field(default_factory=list)
    requested_case_mode: str = "none"
    evidence_required: bool = False
    restart_required: bool = False
    reporting_required: bool = False
    prohibited_claim_types: list[str] = field(default_factory=list)
    allowed_retry_count: int = 1

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class AnswerValidation:
    status: str
    missing_sections: list[str] = field(default_factory=list)
    missing_concepts: list[str] = field(default_factory=list)
    sequence_errors: list[str] = field(default_factory=list)
    unrelated_topic_warnings: list[str] = field(default_factory=list)
    unsupported_number_warnings: list[str] = field(default_factory=list)
    unsupported_legal_reference_warnings: list[str] = field(default_factory=list)
    incomplete_response: bool = False
    incomplete_reason: str = ""
    retry_recommended: bool = False
    fallback_required: bool = False

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class GuardrailResult:
    answer: str
    validation: AnswerValidation
    generation_calls: int
    retry_used: bool
    fallback_used: bool

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["validation"] = self.validation.to_dict()
        return payload


SECTION_TERMS: dict[str, tuple[str, ...]] = {
    "즉시 조치": ("즉시", "우선 조치", "작업중지", "작업을 중지"),
    "조치 순서": ("조치 순서", "우선 조치", "1.", "1)"),
    "점검항목": ("점검", "확인사항", "확인 항목", "상태를 확인", "위치를 확인"),
    "위험 원인": ("왜 위험", "위험 원인", "위험한 이유"),
    "작업중지 여부": ("작업중지", "작업을 중지", "작업 중지", "설비를 정지", "운행을 중지"),
    "대피·출입통제": ("대피", "출입통제", "출입 통제"),
    "작업 재개 조건": ("작업 재개", "재개 조건", "재개 여부", "재개 승인"),
    "보호구": ("보호구", "안전모", "안전대", "마스크"),
    "보고 대상": ("보고", "신고", "책임자"),
    "기록관리": ("기록", "점검표", "승인 내용"),
    "공식 법령·지침 근거": ("관련 근거", "공식 근거", "chunk_id", "근거 문서"),
    "공식 사고사례": ("공식 사고사례", "직접 관련 공식 사례", "유사 위험 사례"),
    "유사 위험 사례": ("유사 위험 사례", "직접 관련 공식 사례"),
    "교육·TBM 공유사항": ("tbm", "교육", "공유사항", "공유 사항"),
    "사업주·관리자 대응": ("관리자", "사업주", "책임자", "지시"),
    "작업 가능 여부 판단": ("작업을 계속", "작업 가능", "작업 재개", "중지"),
    "예방대책": ("예방대책", "개선조치", "재발방지", "감소대책"),
}

CONCEPT_ALIASES: dict[str, tuple[str, ...]] = {
    "우천": ("우천", "강우", "비"),
    "낙뢰": ("낙뢰", "번개"),
    "발파공 침수": ("발파공 침수", "공내수", "발파공 내부 물", "발파 구멍"),
    "화약류 습윤": ("습윤", "젖", "화약류", "뇌관"),
    "불발": ("불발", "터지지 않은", "잔류약"),
    "잔류 화약류": ("잔류 화약", "잔류약"),
    "환기 정지": ("환기 정지", "환기팬", "송풍", "환기설비"),
    "전원 상실": ("정전", "전원 상실"),
    "가스 상승": ("가스", "측정값", "농도", "검지기"),
    "메탄": ("메탄",),
    "산소 부족": ("산소 부족", "산소가 모자", "산소농도"),
    "유해가스": ("유해가스", "가스"),
    "천반": ("천반", "천장"),
    "출수 증가": ("출수", "유출수", "물줄기", "지하수"),
    "지보 변형": ("지보", "받침", "변형", "뒤틀", "휘어"),
    "균열": ("균열",),
    "박리": ("박리",),
    "부석": ("부석",),
    "낙반": ("낙반", "붕락"),
    "전기설비": ("전기설비", "배전", "차단기", "모터", "판넬"),
    "정비": ("정비", "점검", "수리"),
    "전원 차단": ("전원 차단", "전기 차단", "차단기"),
    "잠금·표지": ("잠금", "표지", "loto"),
    "잔류에너지 제거": ("잔류에너지", "잔류 에너지"),
    "무전압": ("무전압", "전압이 없는"),
    "무전압 확인": ("무전압", "전압이 없는"),
    "재투입 방지": ("재투입 방지", "임의 재투입", "재가동 방지"),
    "컨베이어": ("컨베이어", "벨트"),
    "파쇄기": ("파쇄기", "분쇄설비"),
    "운전 중 제거": ("운전 중", "가동 중", "청소", "제거"),
    "끼임": ("끼임", "말림"),
    "에너지 차단": ("에너지 차단", "전원 차단", "잠금"),
    "차량": ("차량", "덤프트럭", "운반장비"),
    "장비": ("장비", "로더", "굴착기"),
    "후진": ("후진", "뒤로"),
    "충돌": ("충돌", "접촉", "스쳤", "칠 뻔"),
    "사각지대": ("사각지대", "보이지"),
    "신호수": ("신호수", "유도자"),
    "분진": ("분진", "먼지"),
    "집진": ("집진",),
    "살수": ("살수",),
    "호흡보호구": ("호흡보호구", "방진마스크", "마스크"),
    "시야 저하": ("시야", "앞이 안 보", "앞이 흐려"),
    "보호구": ("보호구", "안전모", "안전대", "보호장비"),
    "미착용": ("미착용", "착용하지", "안 썼"),
    "파손": ("파손", "깨졌", "찢어진", "손상"),
    "지급": ("지급",),
    "교체": ("교체", "새 보호구"),
    "적합성": ("적합", "맞는 보호구", "안 맞"),
    "위험성평가": ("위험성평가", "위험을 평가"),
    "변경 작업": ("변경", "바뀌", "달라"),
    "작업중지": ("작업중지", "작업 중지", "작업을 중지", "설비를 정지", "운행을 중지", "멈추"),
    "작업 재개": ("작업 재개", "재가동", "재입갱", "재운행"),
    "의식상실": ("의식", "쓰러"),
    "응급조치": ("응급조치", "응급", "부상"),
    "구조": ("구조", "구급"),
    "2차 위험": ("2차 위험", "구조자 안전", "안전이 확보"),
    "사고": ("사고", "부상", "쓰러"),
    "TBM": ("tbm", "티비엠"),
    "굴진": ("굴진", "막장"),
    "천공": ("천공",),
    "작업 전 공유": ("작업 전", "tbm", "공유"),
    "변경사항": ("변경", "바뀌", "달라"),
    "사고보고": ("사고보고", "보고", "신고"),
    "기록": ("기록", "증거", "자료"),
    "현장보존": ("현장보존", "현장 보존"),
    "중대재해": ("중대재해", "중상", "큰 부상"),
    "재발방지": ("재발방지", "재발 방지"),
    "대피·출입통제": ("대피", "출입통제", "출입 통제", "접근을 통제", "접근 통제", "접근하지 않도록 통제"),
    "책임자 보고": ("책임자", "안전관리자", "보고"),
    "안전한 위치에서 점검": ("안전한 위치", "안전한 장소", "안전한 상태", "점검"),
    "위험 제거": ("위험 제거", "위험이 제거", "위험요인 제거", "개선조치"),
    "책임자 확인": ("책임자 확인", "책임자가 확인", "안전관리자 확인", "승인"),
    "작업 재개 판단": ("작업 재개", "재개 여부", "재개 승인"),
    "조치·점검·승인 기록": ("기록", "점검표", "승인 내용"),
    "환기 복구": ("환기 복구", "환기설비 복구", "송풍 복구"),
    "가스·산소 재측정": ("재측정", "가스 측정", "산소 측정"),
    "재개 승인": ("재개 승인", "책임자 확인"),
    "천반·지보·출수·배수 점검": ("천반", "지보", "출수", "배수"),
    "균열·박리·부석 확인": ("균열", "박리", "부석"),
    "배수·집수 검토": ("배수", "집수", "차수"),
    "기상·낙뢰 확인": ("기상", "강우", "우천", "낙뢰"),
    "화약류·뇌관 습윤 확인": ("화약류", "뇌관", "습윤", "젖"),
    "배수·접근로 점검": ("배수", "접근로", "통행로"),
    "작업 연기·중지": ("작업 연기", "작업중지", "작업을 중지"),
}


def _unique(values: Iterable[str]) -> list[str]:
    return list(dict.fromkeys(value for value in values if value))


def _concept_group(concept: str) -> list[str]:
    return list(CONCEPT_ALIASES.get(concept, (concept,)))


def _standard_sequence() -> list[list[str]]:
    return [
        ["작업중지", "작업 중지", "작업을 중지"],
        ["대피", "안전한 장소"],
        ["출입통제", "출입 통제", "위험구역"],
        ["보고", "책임자", "안전관리자"],
        ["안전한 위치", "점검", "측정"],
        ["위험 제거", "개선조치", "보강", "복구"],
        ["작업 재개", "재개 승인", "책임자 확인"],
        ["기록", "점검표", "승인 내용"],
    ]


def build_answer_contract(understanding: QueryUnderstanding) -> AnswerContract:
    outputs = understanding.requested_outputs
    sections = list(outputs)
    simple_education = (
        understanding.work_stage == "교육·회의 단계"
        and understanding.urgency != "high"
        and not {"즉시 조치", "조치 순서", "작업중지 여부"}.intersection(outputs)
    )
    if understanding.ambiguity_level in {"medium", "high"}:
        sections.extend(("즉시 조치", "확인 질문"))

    needs_sequence = "조치 순서" in outputs or understanding.urgency == "high"
    signals = set(understanding.hazard_signals)
    if understanding.primary_domain == DOMAIN_BLASTING and needs_sequence and "발파공 침수" in signals and "불발" not in signals:
        sequence = [
            ["기상", "강우", "낙뢰"],
            ["발파공", "공내수", "배수"],
            ["화약류", "뇌관", "습윤"],
            ["작업 연기", "작업을 중지", "사용하지"],
            ["책임자", "안전관리자", "기록"],
        ]
    elif understanding.primary_domain == DOMAIN_BLASTING and needs_sequence:
        sequence = [
            ["작업중지", "작업을 중지"],
            ["접근", "출입통제", "통제"],
            ["발파책임자", "안전관리자", "보고"],
            ["위험 제거", "현장 확인"],
            ["재개", "책임자 확인"],
            ["기록", "조치 내용"],
        ]
    elif understanding.primary_domain == DOMAIN_VENTILATION and needs_sequence:
        sequence = [
            ["작업중지", "작업을 중지"],
            ["대피", "안전한 장소"],
            ["출입통제", "출입을 통제"],
            ["환기설비", "가스", "산소"],
            ["환기 복구", "환기를 복구"],
            ["재측정", "가스·산소"],
            ["작업 재개", "재개 승인"],
            ["기록", "승인 내용"],
        ]
    elif understanding.primary_domain == DOMAIN_ELECTRICAL and needs_sequence:
        sequence = [
            ["작업중지", "작업을 중지"],
            ["전원 차단", "전원을 차단"],
            ["잠금", "표지"],
            ["잔류에너지", "잔류 에너지"],
            ["무전압", "전압이 없는"],
            ["재투입", "재가동 방지"],
            ["책임자", "재개"],
            ["기록", "점검표"],
        ]
    elif understanding.primary_domain == DOMAIN_CONVEYOR and needs_sequence:
        sequence = [
            ["정지", "작업중지"],
            ["접근", "통제"],
            ["에너지", "전원", "잠금"],
            ["재가동", "재투입"],
            ["제거", "방호장치"],
            ["책임자", "인원 확인"],
            ["기록", "승인"],
        ]
    elif understanding.primary_domain == DOMAIN_MOBILE and needs_sequence:
        sequence = [
            ["운행을 중지", "이동을 중지"],
            ["작업자", "분리", "통제"],
            ["신호수", "사각지대", "경고장치"],
            ["위험요인 제거", "개선조치"],
            ["운행 재개", "책임자"],
            ["기록", "승인"],
        ]
    elif understanding.primary_domain == DOMAIN_DUST and needs_sequence:
        sequence = [
            ["작업을 중지", "작업중지"],
            ["집진", "살수", "환기"],
            ["호흡보호구", "마스크"],
            ["책임자", "재개"],
            ["기록", "점검"],
        ]
    elif understanding.primary_domain == DOMAIN_PPE:
        sequence = [
            ["작업 위험", "위험요인"],
            ["보호구 선정", "적합"],
            ["지급", "착용"],
            ["파손", "교체"],
            ["교육", "기록"],
        ] if "조치 순서" in outputs else []
    elif understanding.primary_domain == DOMAIN_REPORTING and understanding.work_stage == "사고 발생 직후":
        sequence = [
            ["응급조치", "구조"],
            ["현장보존", "현장 보존", "통제"],
            ["보고", "신고"],
            ["기록", "자료"],
        ]
    elif understanding.primary_domain == DOMAIN_EMERGENCY:
        sequence = [
            ["2차 위험", "구조자 안전", "안전 확인"],
            ["작업중지", "접근통제", "출입 통제"],
            ["구조", "응급조치"],
            ["보고", "신고"],
            ["기록", "현장보존"],
        ]
    elif not simple_education and needs_sequence:
        sequence = _standard_sequence()
    else:
        sequence = []

    concepts = [_concept_group(concept) for concept in understanding.required_concepts]
    evidence_required = "공식 법령·지침 근거" in outputs
    case_mode = "none"
    if "공식 사고사례" in outputs:
        case_mode = "official"
    elif "유사 위험 사례" in outputs:
        case_mode = "similar"
    return AnswerContract(
        required_sections=_unique(sections),
        required_concept_groups=concepts,
        required_sequence=sequence,
        requested_case_mode=case_mode,
        evidence_required=evidence_required,
        restart_required=bool({"작업 재개 조건", "작업 가능 여부 판단"}.intersection(outputs)),
        reporting_required=bool({"보고 대상", "기록관리", "사업주·관리자 대응"}.intersection(outputs)),
        prohibited_claim_types=[
            "검색 근거에 없는 중요 수치",
            "검색 근거에 없는 법령 조항",
            "법령 위반·처벌 확정 표현",
        ],
        allowed_retry_count=1,
    )


def incomplete_response_reason(answer: str, finish_reason: str = "") -> str:
    text = str(answer or "").strip()
    finish = str(finish_reason or "").upper()
    if any(marker in finish for marker in ("MAX_TOKENS", "LENGTH", "SAFETY", "MALFORMED")):
        return f"비정상 종료 사유: {finish}"
    if not text:
        return "빈 응답"
    if len(re.sub(r"\s+", "", text)) < 35:
        return "지나치게 짧은 응답"
    if any(fragment in text for fragment in ("제공된 [안정", "[반드시 아래 구조로 출력]", "아래 [안정형 조치 초안]")):
        return "프롬프트 조각 노출"
    if text.count("[") > text.count("]"):
        return "닫히지 않은 대괄호"
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    last = lines[-1] if lines else ""
    if last.endswith((":", ",", "，", ";", "그리고", "그러나", "따라서", "때문에", "또는", " 및", "-", "•", "→")):
        return "문장 또는 목록의 갑작스러운 종료"
    if last.startswith(("-", "*", "•")) and len(last.lstrip("-*• ").strip()) < 3:
        return "비어 있는 목록 항목"
    return ""


def _group_present(text: str, group: Iterable[str]) -> bool:
    lowered = text.lower()
    return any(str(term).lower() in lowered for term in group)


def _sequence_errors(text: str, sequence: list[list[str]]) -> list[str]:
    if not sequence:
        return []
    cursor = 0
    errors: list[str] = []
    for index, group in enumerate(sequence, start=1):
        positions = [text.find(term, cursor) for term in group if text.find(term, cursor) >= 0]
        if not positions:
            errors.append(f"{index}단계 누락 또는 순서 불일치: {'/'.join(group[:3])}")
            continue
        cursor = min(positions) + 1
    return errors


def _evidence_corpus(
    understanding: QueryUnderstanding,
    evidence_results: list[dict[str, Any]],
) -> str:
    parts = [understanding.normalized_query]
    for result in evidence_results:
        parts.extend((
            str(result.get("source", "")),
            str(result.get("chunk_id", "")),
            str(result.get("text", "")),
        ))
    return "\n".join(parts)


def _unsupported_numbers(answer: str, corpus: str) -> list[str]:
    pattern = re.compile(
        r"(?<![\w.])(\d+(?:\.\d+)?)\s*(%|ppm|mg/m3|mg/㎥|분|시간|미터|m|회|만원|원|년 이하|개월)"
    )
    warnings: list[str] = []
    for match in pattern.finditer(answer):
        claim = match.group(0)
        if claim not in corpus and match.group(1) not in corpus:
            warnings.append(claim)
    return _unique(warnings)


def _unsupported_legal_references(answer: str, corpus: str) -> list[str]:
    references = re.findall(r"제\s*\d+\s*조(?:의\s*\d+)?", answer)
    return _unique(reference for reference in references if reference not in corpus)


def official_case_status_message(
    understanding: QueryUnderstanding,
    official_cases: list[dict[str, Any]],
) -> str:
    if not understanding.official_case_requested:
        return ""
    if not official_cases:
        return (
            "현재 품질검사를 통과한 직접 관련 공식 사례는 확인되지 않았습니다. "
            "무관한 사례는 표시하지 않습니다."
        )
    direct = [case for case in official_cases if str(case.get("relation", "direct")).lower() == "direct"]
    if direct:
        return "품질검사를 통과한 직접 관련 공식 사고사례를 별도 사례 영역에 표시합니다."
    return "직접 관련 사례가 없어 검증된 유사 위험 사례로 구분해 표시합니다."


def validate_answer(
    answer: str,
    contract: AnswerContract,
    understanding: QueryUnderstanding,
    *,
    evidence_results: list[dict[str, Any]] | None = None,
    official_cases: list[dict[str, Any]] | None = None,
    finish_reason: str = "",
) -> AnswerValidation:
    text = str(answer or "").strip()
    evidence_results = evidence_results or []
    official_cases = official_cases or []
    incomplete_reason = incomplete_response_reason(text, finish_reason)

    missing_sections: list[str] = []
    for section in contract.required_sections:
        if section == "확인 질문":
            present = bool(understanding.clarification_question) and (
                understanding.clarification_question in text or "확인 질문" in text
            )
        else:
            present = _group_present(text, SECTION_TERMS.get(section, (section,)))
        if not present:
            missing_sections.append(section)

    missing_concepts = [
        "/".join(group[:4])
        for group in contract.required_concept_groups
        if not _group_present(text, group)
    ]
    sequence_errors = _sequence_errors(text, contract.required_sequence)

    unrelated: list[str] = []
    relevant_hits = sum(
        1 for group in contract.required_concept_groups if _group_present(text, group)
    )
    for topic in understanding.forbidden_dominant_topics:
        count = text.lower().count(topic.lower())
        if count >= 2 and count > max(1, relevant_hits):
            unrelated.append(f"질문과 무관한 주제가 과다함: {topic}")

    corpus = _evidence_corpus(understanding, evidence_results)
    number_warnings = _unsupported_numbers(text, corpus)
    legal_warnings = _unsupported_legal_references(text, corpus)
    if contract.evidence_required:
        complete_evidence = [
            result for result in evidence_results
            if str(result.get("source", "")).strip() and str(result.get("chunk_id", "")).strip()
        ]
        cited = any(
            str(result.get("source", "")) in text
            and str(result.get("chunk_id", "")) in text
            for result in complete_evidence
        )
        if not complete_evidence or not cited:
            missing_sections.append("실제 문서명·chunk_id 근거")

    if contract.requested_case_mode != "none":
        case_message = official_case_status_message(understanding, official_cases)
        case_fields = [
            str(case.get("display_title") or case.get("title") or "")
            for case in official_cases
        ]
        if case_message not in text and not any(field and field in text for field in case_fields):
            missing_sections.append("공식 사고사례 처리 상태")

    missing_sections = _unique(missing_sections)
    issues = any((
        missing_sections,
        missing_concepts,
        sequence_errors,
        unrelated,
        number_warnings,
        legal_warnings,
        incomplete_reason,
    ))
    status = "질문 핵심 요소 반영 충분" if not issues else "질문 핵심 요소 반영 보완 필요"
    return AnswerValidation(
        status=status,
        missing_sections=missing_sections,
        missing_concepts=missing_concepts,
        sequence_errors=sequence_errors,
        unrelated_topic_warnings=unrelated,
        unsupported_number_warnings=number_warnings,
        unsupported_legal_reference_warnings=legal_warnings,
        incomplete_response=bool(incomplete_reason),
        incomplete_reason=incomplete_reason,
        retry_recommended=issues,
        fallback_required=bool(incomplete_reason and not text),
    )


def format_answer_contract_prompt(contract: AnswerContract) -> str:
    sections = ", ".join(contract.required_sections) or "질문에 직접 답변"
    concepts = ", ".join(
        group[0] for group in contract.required_concept_groups if group
    ) or "질문의 세부 위험 신호"
    sequence = " → ".join(group[0] for group in contract.required_sequence if group)
    lines = [
        f"- 필수 답변 항목: {sections}",
        f"- 반드시 보존할 개념: {concepts}",
    ]
    if sequence:
        lines.append(f"- 필요한 조치 순서: {sequence}")
    if contract.evidence_required:
        lines.append("- 공식 근거는 검색 결과의 실제 문서명과 chunk_id만 사용")
    if contract.requested_case_mode != "none":
        lines.append("- 사례 요청 처리 상태를 명시하고 무관한 사례는 표시하지 않음")
    lines.append("- 검색 근거에 없는 수치·시간·법령 조항·처벌을 만들지 않음")
    return "\n".join(lines)


def build_complete_core_judgment(understanding: QueryUnderstanding) -> str:
    if understanding.ambiguity_level in {"medium", "high"}:
        return (
            "위험의 정확한 원인이 확인되지 않았으므로 해당 구역의 작업을 중지하고 접근을 통제해야 합니다. "
            "안전한 위치에서 핵심 발생 위치와 상태를 확인하기 전에는 작업 재개를 확정하지 않습니다."
        )
    judgments = {
        DOMAIN_BLASTING: "발파 단계와 기상·발파공·화약류 상태를 구분해 확인하고 위험이 남아 있으면 작업을 중지해야 합니다.",
        DOMAIN_VENTILATION: "환기 또는 가스 이상은 작업중지와 대피를 우선하고 환기 복구와 재측정 전에는 재개하지 않아야 합니다.",
        DOMAIN_GROUND: "천반·지보 이상은 낙반 전조로 보고 작업중지·대피·출입통제를 우선해야 합니다.",
        DOMAIN_ELECTRICAL: "전기설비 작업은 에너지를 차단하고 잠금·표지와 무전압 확인을 마치기 전에는 시작하지 않아야 합니다.",
        DOMAIN_CONVEYOR: "가동 중인 컨베이어·파쇄설비에는 접근하지 말고 정지와 에너지 차단을 먼저 해야 합니다.",
        DOMAIN_MOBILE: "차량·장비의 이동 구역은 작업자와 분리하고 사각지대·신호 상태가 불명확하면 운행을 중지해야 합니다.",
        DOMAIN_DUST: "분진이 제어되지 않으면 작업을 중지하고 집진·살수·환기와 호흡보호 상태를 확인해야 합니다.",
        DOMAIN_PPE: "작업 위험에 적합하고 손상되지 않은 보호구를 지급·착용하기 전에는 작업을 시작하지 않아야 합니다.",
        DOMAIN_RISK: "위험이 불명확하거나 제거되지 않았으면 작업을 중지하고 책임자 확인 전에는 재개하지 않아야 합니다.",
        DOMAIN_EMERGENCY: "구조자의 2차 위험을 먼저 통제한 뒤 가능한 범위에서 구조·응급조치와 보고를 실시해야 합니다.",
        DOMAIN_TBM: "작업·장비·장소의 변경 위험을 TBM에서 공유하고 조치가 확인된 뒤 작업을 시작해야 합니다.",
        DOMAIN_REPORTING: "인명 보호와 현장 통제를 우선한 뒤 현장보존·보고·기록을 사실에 따라 수행해야 합니다.",
    }
    return judgments.get(understanding.primary_domain, "위험을 확인하고 보수적인 현장조치를 우선해야 합니다.")


def _domain_actions(understanding: QueryUnderstanding) -> list[str]:
    signals = set(understanding.hazard_signals)
    domain = understanding.primary_domain
    if domain == DOMAIN_BLASTING and {"우천", "발파공 침수"}.intersection(signals):
        return [
            "기상과 낙뢰 상태를 확인하고 위험하면 발파 작업을 연기하거나 중지합니다.",
            "발파공 내부 물·공내수와 작업장 배수·접근로 상태를 확인합니다.",
            "화약류·뇌관·점화설비의 습윤·손상을 확인하고 이상품을 사용하지 않습니다.",
            "사면·갱구·천반·부석 상태를 확인하고 책임자가 점검 결과와 작업 판단을 기록합니다.",
        ]
    if domain == DOMAIN_BLASTING and "불발" in signals:
        return [
            "불발 의심 구역의 작업을 중지하고 접근·출입을 통제합니다.",
            "발파책임자 또는 안전관리자에게 보고하고 임의로 불발공을 처리하지 않습니다.",
            "현장 위험 제거와 책임자 확인 후 재개 여부를 판단하고 조치 내용을 기록합니다.",
        ]
    if domain == DOMAIN_VENTILATION:
        return [
            "작업을 중지하고 작업자를 안전한 장소로 대피시킨 뒤 출입을 통제합니다.",
            "안전한 위치에서 정전·전원 상실 원인, 환기설비와 가스·산소 상태를 점검합니다.",
            "환기를 복구하고 가스·산소를 재측정한 뒤 책임자가 작업 재개 여부를 판단합니다.",
            "이상 발견·대피·측정·복구·재개 승인 내용을 기록합니다.",
        ]
    if domain == DOMAIN_GROUND:
        return [
            "해당 구간 작업을 중지하고 작업자를 안전한 장소로 대피시킨 뒤 출입을 통제합니다.",
            "책임자에게 보고하고 안전한 위치에서 천반·지보·출수·배수 상태를 점검합니다.",
            "균열·박리·부석과 지보재 변형을 확인하고 배수·집수 또는 필요한 보강을 검토합니다.",
            "위험 제거와 책임자 확인 후 재개 여부를 판단하고 모든 조치와 승인을 기록합니다.",
        ]
    if domain == DOMAIN_ELECTRICAL:
        return [
            "작업을 중지하고 전원을 차단한 뒤 위험구역 접근을 통제합니다.",
            "잠금·표지를 하고 잔류에너지를 제거한 뒤 무전압을 확인합니다.",
            "임의 재투입을 방지하고 이상을 제거한 뒤 책임자가 재개 여부를 확인합니다.",
            "차단·점검·조치·재개 승인 내용을 기록합니다.",
        ]
    if domain == DOMAIN_CONVEYOR:
        return [
            "설비를 정지하고 작업자가 가동부에 접근하지 않도록 통제합니다.",
            "전원과 잔류에너지를 차단하고 잠금·표지로 재가동을 방지합니다.",
            "막힘·끼임 원인을 안전한 상태에서 제거하고 방호장치와 작업자 위치를 확인합니다.",
            "책임자 확인과 인원 확인 후 재가동 여부를 판단하고 기록합니다.",
        ]
    if domain == DOMAIN_MOBILE:
        return [
            "차량·장비 이동을 중지하고 작업자와 운행 구역을 분리합니다.",
            "신호수·유도자 위치, 사각지대, 경고장치와 통행 상태를 점검합니다.",
            "위험요인을 제거하고 책임자가 운행 재개 여부를 확인합니다.",
            "아차사고·점검·지시·재운행 승인 내용을 기록합니다.",
        ]
    if domain == DOMAIN_DUST:
        return [
            "분진이 제어되지 않거나 시야가 확보되지 않으면 작업을 중지합니다.",
            "집진·살수·환기 상태와 분진 발생원을 점검합니다.",
            "작업 위험에 적합한 호흡보호구의 착용·밀착·손상 상태를 확인합니다.",
            "분진 제어와 책임자 확인 후 재개 여부를 판단하고 기록합니다.",
        ]
    if domain == DOMAIN_PPE:
        return [
            "작업 위험을 확인하고 그 위험에 적합한 보호구를 선정합니다.",
            "보호구를 지급하고 올바른 착용과 적합성을 확인합니다.",
            "파손·오염·부적합 보호구는 사용을 중지하고 교체합니다.",
            "지급·착용점검·교체·교육 내용을 기록합니다.",
        ]
    if domain == DOMAIN_EMERGENCY:
        return [
            "구조 장소의 가스·전기·낙반 등 2차 위험을 확인하고 접근을 통제합니다.",
            "구조자 안전을 확보한 범위에서 구조와 응급조치를 실시하고 필요한 신고를 합니다.",
            "책임자에게 보고하고 현장을 보존하며 부상 상태와 조치 내용을 기록합니다.",
        ]
    if domain == DOMAIN_TBM:
        return [
            "작업 전 작업 위치·방법·장비의 변경사항과 위험요인을 확인합니다.",
            "TBM에서 역할·신호·대피·작업중지 기준과 예방조치를 공유합니다.",
            "작업자의 이해와 보호구·장비 상태를 확인하고 참석·지시·조치 내용을 기록합니다.",
        ]
    if domain == DOMAIN_REPORTING:
        return [
            "인명 보호를 위한 구조·응급조치와 추가 위험 통제를 우선합니다.",
            "안전한 범위에서 현장을 보존하고 정해진 보고체계에 따라 책임자에게 보고합니다.",
            "발생 사실·피해·조치·현장 자료를 기록하고 재발방지대책과 이행을 관리합니다.",
            "위험 제거와 책임자 확인 후 작업 재개 여부를 판단하고 승인 내용을 기록합니다.",
        ]
    return [
        "해당 구역 작업을 중지하고 위험구역 접근을 통제합니다.",
        "안전한 위치에서 위험의 발생 위치와 대상을 확인하고 책임자에게 보고합니다.",
        "위험 제거와 책임자 확인 전에는 작업 재개를 확정하지 않고 조치 내용을 기록합니다.",
    ]


def build_rule_based_fallback(
    understanding: QueryUnderstanding,
    contract: AnswerContract | None = None,
    *,
    evidence_results: list[dict[str, Any]] | None = None,
    official_cases: list[dict[str, Any]] | None = None,
) -> str:
    contract = contract or build_answer_contract(understanding)
    evidence_results = evidence_results or []
    official_cases = official_cases or []
    lines = [
        "## 핵심 판단",
        build_complete_core_judgment(understanding),
        "",
        "## 우선 조치",
    ]
    lines.extend(
        f"{index}. {action}"
        for index, action in enumerate(_domain_actions(understanding), start=1)
    )
    if understanding.clarification_question:
        lines.extend((
            "",
            "## 확인 질문",
            understanding.clarification_question,
            "확인 전에는 특정 원인이나 작업 재개 가능 여부를 확정하지 않습니다.",
        ))
    if contract.restart_required:
        lines.extend((
            "",
            "## 작업 재개 조건",
            "위험이 제거되고 필요한 점검·측정이 끝난 뒤 책임자가 확인한 경우에만 작업 재개 여부를 판단합니다.",
        ))
    if contract.reporting_required:
        lines.extend((
            "",
            "## 보고 및 기록관리",
            "발견 시점, 작업중지·대피·통제, 점검 결과, 개선조치와 재개 승인 내용을 사실대로 기록합니다.",
        ))
    if contract.evidence_required:
        lines.extend(("", "## 관련 근거 문서"))
        complete = [
            result for result in evidence_results
            if str(result.get("source", "")).strip() and str(result.get("chunk_id", "")).strip()
        ]
        if complete:
            for result in complete[:3]:
                lines.append(f"- {result['source']} (chunk_id: {result['chunk_id']})")
        else:
            lines.append("- 공식 근거 보완이 필요합니다. 확인되지 않은 법령 조항이나 수치를 만들지 않습니다.")
            lines.append("- 현장 안전관리자·관계기관 또는 전문가에게 최신 기준을 확인하세요.")
    case_status = official_case_status_message(understanding, official_cases)
    if case_status:
        lines.extend(("", "## 공식 사고사례", case_status))
    return "\n".join(lines).strip()


def guard_answer_generation(
    primary_answer: str,
    understanding: QueryUnderstanding,
    contract: AnswerContract,
    *,
    regenerate: Callable[[AnswerValidation], str] | None = None,
    fallback_factory: Callable[[], str] | None = None,
    evidence_results: list[dict[str, Any]] | None = None,
    official_cases: list[dict[str, Any]] | None = None,
    finish_reason: str = "",
) -> GuardrailResult:
    validation = validate_answer(
        primary_answer,
        contract,
        understanding,
        evidence_results=evidence_results,
        official_cases=official_cases,
        finish_reason=finish_reason,
    )
    if not validation.retry_recommended:
        return GuardrailResult(primary_answer, validation, 1, False, False)

    calls = 1
    if regenerate is not None and contract.allowed_retry_count > 0:
        retry_answer = str(regenerate(validation) or "")
        calls += 1
        retry_validation = validate_answer(
            retry_answer,
            contract,
            understanding,
            evidence_results=evidence_results,
            official_cases=official_cases,
        )
        if not retry_validation.retry_recommended:
            return GuardrailResult(retry_answer, retry_validation, calls, True, False)
        validation = retry_validation

    fallback = (
        str(fallback_factory() or "")
        if fallback_factory is not None
        else build_rule_based_fallback(
            understanding,
            contract,
            evidence_results=evidence_results,
            official_cases=official_cases,
        )
    )
    fallback_validation = validate_answer(
        fallback,
        contract,
        understanding,
        evidence_results=evidence_results,
        official_cases=official_cases,
    )
    return GuardrailResult(fallback, fallback_validation, calls, calls > 1, True)


__all__ = [
    "AnswerContract",
    "AnswerValidation",
    "GuardrailResult",
    "build_answer_contract",
    "build_complete_core_judgment",
    "build_rule_based_fallback",
    "format_answer_contract_prompt",
    "guard_answer_generation",
    "incomplete_response_reason",
    "official_case_status_message",
    "validate_answer",
]
