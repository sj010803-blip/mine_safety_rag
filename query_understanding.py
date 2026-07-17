from __future__ import annotations

from dataclasses import asdict, dataclass, field
import re
from typing import Any, Iterable


DOMAIN_BLASTING = "발파·불발·화약류"
DOMAIN_VENTILATION = "환기·메탄·유해가스·산소부족"
DOMAIN_GROUND = "낙반·붕괴·천반·지보"
DOMAIN_ELECTRICAL = "전기·감전·LOTO"
DOMAIN_CONVEYOR = "컨베이어·끼임·파쇄"
DOMAIN_MOBILE = "차량·장비·후진·충돌"
DOMAIN_DUST = "분진·집진·살수·호흡보호"
DOMAIN_PPE = "보호구 지급·파손·미착용"
DOMAIN_RISK = "위험성평가·작업중지·작업재개"
DOMAIN_EMERGENCY = "응급조치·의식상실·구조"
DOMAIN_TBM = "TBM·굴진·천공"
DOMAIN_REPORTING = "사고보고·기록·중대재해 대응"
OUT_OF_SCOPE_DOMAIN = "범위 밖"

SUPPORTED_DOMAINS = (
    DOMAIN_BLASTING,
    DOMAIN_VENTILATION,
    DOMAIN_GROUND,
    DOMAIN_ELECTRICAL,
    DOMAIN_CONVEYOR,
    DOMAIN_MOBILE,
    DOMAIN_DUST,
    DOMAIN_PPE,
    DOMAIN_RISK,
    DOMAIN_EMERGENCY,
    DOMAIN_TBM,
    DOMAIN_REPORTING,
)

TYPO_NORMALIZATIONS = {
    "대처방안 수입": "대처방안 수립",
    "광산안전전관리자": "광산안전관리자",
}

DOMAIN_RULES: dict[str, tuple[tuple[str, int], ...]] = {
    DOMAIN_BLASTING: (
        ("불발", 6), ("잔류약", 6), ("잔류 화약", 6), ("발파", 5),
        ("장약", 4), ("폭약", 4), ("화약", 4), ("뇌관", 4),
    ),
    DOMAIN_VENTILATION: (
        ("유해가스", 6), ("환기", 5), ("송풍", 5), ("메탄", 5),
        ("산소", 4), ("가스 측정", 4), ("검지기", 3), ("환기팬", 5),
        ("밀폐", 3), ("공기가", 2), ("가스 냄새", 5), ("무슨 가스", 5),
        ("가스", 2),
    ),
    DOMAIN_GROUND: (
        ("낙반", 6), ("붕락", 6), ("천반", 5), ("지보", 5),
        ("부석", 5), ("유출수", 4), ("출수", 4), ("천장", 4),
        ("받침", 3), ("지하수", 3), ("균열", 2), ("박리", 3),
    ),
    DOMAIN_ELECTRICAL: (
        ("loto", 7), ("무전압", 6), ("잠금표지", 6), ("잠금 표지", 6),
        ("감전", 6), ("누전", 6), ("전원 차단", 5), ("전원차단", 5),
        ("배전", 4), ("전기", 4), ("차단기", 3), ("판넬", 3),
        ("정전", 3), ("모터", 2),
    ),
    DOMAIN_CONVEYOR: (
        ("컨베이어", 7), ("파쇄기", 7), ("분쇄설비", 7), ("벨트", 5),
        ("끼임", 5), ("롤러", 4), ("막힘 제거", 4), ("낀 돌", 4),
    ),
    DOMAIN_MOBILE: (
        ("덤프트럭", 7), ("운반장비", 6), ("굴착기", 6), ("로더", 6),
        ("후진", 5), ("충돌", 5), ("신호수", 5), ("유도자", 5),
        ("사각지대", 5), ("운반차량", 5), ("칠 뻔", 5), ("스쳤", 5),
        ("차량", 4), ("장비", 2),
    ),
    DOMAIN_DUST: (
        ("분진", 7), ("집진", 7), ("살수", 6), ("먼지", 5),
        ("호흡보호구", 5), ("방진마스크", 5), ("시야", 2),
    ),
    DOMAIN_PPE: (
        ("보호구", 7), ("안전모", 7), ("안전대", 7), ("작업화", 6),
        ("보안경", 6), ("귀마개", 6), ("보호장비", 6), ("미착용", 4),
        ("마스크", 3),
    ),
    DOMAIN_RISK: (
        ("위험성평가", 8), ("작업중지", 6), ("작업 중지", 6),
        ("작업재개", 6), ("작업 재개", 6), ("재가동", 4),
        ("공정 변경", 5), ("새 작업", 4), ("계속 작업", 3),
    ),
    DOMAIN_EMERGENCY: (
        ("의식", 7), ("쓰러", 7), ("응급", 7), ("구조", 6),
        ("2차 위험", 5), ("부상", 5), ("다쳤", 5), ("119", 6),
        ("구조자", 5),
    ),
    DOMAIN_TBM: (
        ("tbm", 8), ("티비엠", 8), ("굴진", 6), ("천공", 6),
        ("막장", 5),
    ),
    DOMAIN_REPORTING: (
        ("사고보고", 8), ("사고 보고", 8), ("중대재해", 8),
        ("현장보존", 7), ("현장 보존", 7), ("재발방지", 7),
        ("아차사고", 7), ("증거", 4), ("기록", 3), ("보고", 2),
        ("중상 사고", 6), ("큰 부상", 5), ("법 위반", 3), ("사고", 1),
    ),
}

DOMAIN_ORDER = {domain: index for index, domain in enumerate(SUPPORTED_DOMAINS)}


@dataclass(frozen=True)
class QueryUnderstanding:
    original_query: str
    normalized_query: str
    primary_domain: str
    secondary_domains: list[str] = field(default_factory=list)
    hazard_signals: list[str] = field(default_factory=list)
    work_stage: str = "불명확"
    actor: str = "불명확"
    requested_outputs: list[str] = field(default_factory=list)
    urgency: str = "normal"
    ambiguity_level: str = "low"
    ambiguity_reasons: list[str] = field(default_factory=list)
    clarification_question: str = ""
    search_queries: list[str] = field(default_factory=list)
    case_search_query: str = ""
    required_concepts: list[str] = field(default_factory=list)
    forbidden_dominant_topics: list[str] = field(default_factory=list)
    display_label: str = ""
    official_case_requested: bool = False
    in_scope: bool = True

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def normalize_query(query: str) -> str:
    normalized = str(query or "").strip()
    for wrong, corrected in TYPO_NORMALIZATIONS.items():
        normalized = normalized.replace(wrong, corrected)
    return re.sub(r"\s+", " ", normalized)


def _contains_any(text: str, terms: Iterable[str]) -> bool:
    return any(term in text for term in terms)


def _unique(values: Iterable[str]) -> list[str]:
    return list(dict.fromkeys(value for value in values if value))


def _domain_scores(text: str) -> tuple[dict[str, int], dict[str, list[str]]]:
    scores: dict[str, int] = {}
    hits: dict[str, list[str]] = {}
    for domain, rules in DOMAIN_RULES.items():
        matched = [(term, weight) for term, weight in rules if term in text]
        scores[domain] = sum(weight for _, weight in matched)
        hits[domain] = [term for term, _ in matched]
    if "감전" in text and _contains_any(text, ("차단", "전원", "전기")):
        scores[DOMAIN_ELECTRICAL] += 2
    if "감전 사고" in text:
        scores[DOMAIN_ELECTRICAL] += 6
    if "정비" in text and _contains_any(text, ("전기", "전원", "배전", "모터")):
        scores[DOMAIN_ELECTRICAL] += 4
    if _contains_any(text, ("현장보존", "현장 보존", "중대재해", "아차사고")):
        scores[DOMAIN_REPORTING] += 3
    if _contains_any(text, ("2차 위험", "의식을 잃", "쓰러")):
        scores[DOMAIN_EMERGENCY] += 3
    other_hazard_score = max(
        (score for domain, score in scores.items() if domain != DOMAIN_TBM),
        default=0,
    )
    if scores[DOMAIN_TBM] and other_hazard_score >= 4 and _contains_any(
        text, ("공유", "교육", "사례", "예방수칙", "예방대책")
    ):
        scores[DOMAIN_TBM] = max(0, scores[DOMAIN_TBM] - 5)
    return scores, hits


def _is_in_scope(text: str, scores: dict[str, int]) -> bool:
    if max(scores.values(), default=0) >= 3:
        return True
    safety_context = (
        "광산", "갱내", "작업", "위험", "안전", "관리자", "대피", "출입통제",
        "재개", "점검", "정비", "사고", "현장", "들어가도", "사람 빼",
    )
    outside = (
        "맛집", "메뉴", "영화", "여행", "날씨", "주식", "번역", "축구", "요리",
    )
    ambiguous_water_safety_question = (
        _contains_any(text, ("물이", "물은", "물이 새", "물 유입"))
        and _contains_any(text, ("늘", "많아", "증가", "갑자기"))
        and _contains_any(text, ("어떻게 해야", "계속", "들어가도", "작업", "위험"))
    )
    return (_contains_any(text, safety_context) or ambiguous_water_safety_question) and not (
        _contains_any(text, outside) and not _contains_any(text, ("작업", "안전", "광산"))
    )


def _classify_domains(text: str) -> tuple[str, list[str], dict[str, int]]:
    scores, hits = _domain_scores(text)
    if not _is_in_scope(text, scores):
        return OUT_OF_SCOPE_DOMAIN, [], scores
    ranked = sorted(
        SUPPORTED_DOMAINS,
        key=lambda domain: (-scores[domain], DOMAIN_ORDER[domain]),
    )
    primary = ranked[0]
    if scores[primary] < 3:
        primary = DOMAIN_RISK
    primary_score = max(scores.get(primary, 0), 3)
    secondary = [
        domain
        for domain in ranked
        if domain != primary
        and scores[domain] >= 3
        and hits[domain]
    ]
    return primary, secondary[:3], scores


def _extract_signals(text: str) -> list[str]:
    signals: list[str] = []
    has = lambda *terms: _contains_any(text, terms)
    if has("발파", "장약", "폭약", "화약", "뇌관", "불발", "잔류약"):
        signals.append("발파")
    if has("우천", "강우", "비 온", "비온", "비 그친", "비가"):
        signals.append("우천")
    if has("낙뢰", "번개"):
        signals.append("낙뢰")
    if has("발파공", "발파 구멍") and has("물", "침수", "잠긴", "공내수"):
        signals.append("발파공 침수")
    if has("화약", "폭약", "뇌관", "장약", "점화") and has("젖", "습윤", "물", "침수"):
        signals.append("화약류 습윤")
    if has("불발", "터지지 않", "잔류약", "잔류 화약"):
        signals.append("불발")
    if has("잔류약", "잔류 화약", "터지지 않은 약"):
        signals.append("잔류 화약류")

    if has("환기", "송풍", "환기팬", "팬") and has("멈", "정지", "끊", "고장"):
        signals.append("환기 정지")
    elif has("환기설비", "환기 설비") and has("점검", "정비"):
        signals.append("환기 정지")
    if has("정전", "전원 상실", "전원이 끊"):
        signals.append("전원 상실")
    if has("가스", "메탄", "검지기", "측정값", "농도") and has("오르", "올라", "높아", "상승", "증가"):
        signals.append("가스 상승")
    if "메탄" in text:
        signals.append("메탄")
    if has("산소 부족", "산소가 모자", "산소농도 저하"):
        signals.append("산소 부족")
    if has("유해가스", "가스 냄새", "무슨 가스", "공기가 이상"):
        signals.append("유해가스")

    if has("천반", "천장"):
        signals.append("천반")
    roof_context = has("천반", "천장", "측벽", "지보", "출수", "유출수", "지하수")
    if roof_context and has("출수", "유출수", "물", "물줄기", "지하수") and has("늘", "증가", "많", "굵어", "급증"):
        signals.append("출수 증가")
    if has("지보", "받침") and has("휘", "변형", "뒤틀", "손상", "보강"):
        signals.append("지보 변형")
    if "균열" in text:
        signals.append("균열")
    if "박리" in text:
        signals.append("박리")
    if "부석" in text:
        signals.append("부석")
    if has("낙반", "붕락", "떨어질"):
        signals.append("낙반")

    if has("전기", "배전", "판넬", "차단기", "모터", "전원", "감전", "누전"):
        signals.append("전기설비")
    if has("정비", "수리", "막힘 제거"):
        signals.append("정비")
    if has("전원 차단", "전원차단", "전기 차단", "차단기만 내리", "전원부터 끊"):
        signals.append("전원 차단")
    if has("loto", "잠금표지", "잠금 표지", "잠그고 표지"):
        signals.append("잠금·표지")
    if has("무전압", "전기가 없는지"):
        signals.append("무전압")
    if "누전" in text:
        signals.append("누전")
    if "감전" in text:
        signals.append("감전")

    if has("컨베이어", "벨트"):
        signals.append("컨베이어")
    if has("파쇄기", "분쇄설비"):
        signals.append("파쇄기")
    if (
        has("도는 중", "운행 중", "돌고", "운전 중")
        or (has("컨베이어", "벨트") and has("잠깐만", "빨리"))
    ) and has("빼", "제거", "청소", "밀어"):
        signals.append("운전 중 제거")
    if has("끼임", "낀 돌", "걸렸"):
        signals.append("끼임")
    if has("에너지 차단", "재가동 방지") or (
        has("파쇄기", "분쇄설비", "컨베이어", "벨트") and has("차단")
    ):
        signals.append("에너지 차단")

    if has("차량", "덤프트럭", "운반차량", "운반장비"):
        signals.append("차량")
    if has("장비", "굴착기", "로더"):
        signals.append("장비")
    if "후진" in text or "뒤로 갈" in text:
        signals.append("후진")
    if has("충돌", "칠 뻔", "스쳤"):
        signals.append("충돌")
    if "사각지대" in text:
        signals.append("사각지대")
    elif has("신호수", "유도자") and has("안 보여", "보이지"):
        signals.append("사각지대")
    if has("신호수", "유도자"):
        signals.append("신호수")

    if has("분진", "먼지") or ("집진" in text and "천공" in text):
        signals.append("분진")
    if "집진" in text:
        signals.append("집진")
    if "살수" in text:
        signals.append("살수")
    if has("호흡보호구", "방진마스크", "마스크") and has("분진", "먼지", "집진", "호흡"):
        signals.append("호흡보호구")
    if has("앞이 안 보", "앞이 흐려", "시야"):
        signals.append("시야 저하")

    if has("보호구", "안전모", "안전대", "보호장비", "작업화", "보안경", "마스크"):
        signals.append("보호구")
    if has("미착용", "안 썼", "안 찼", "착용 안"):
        signals.append("미착용")
    if has("깨졌", "찢어진", "파손", "손상") and has("보호구", "안전모", "안전대", "보호장비"):
        signals.append("파손")
    if "지급" in text:
        signals.append("지급")
    if "교체" in text or (has("새 보호구", "새 안전모", "새 안전대") and has("전", "지급")):
        signals.append("교체")
    if has("적합", "안 맞", "작업에 맞는") and has("보호구", "보호장비"):
        signals.append("적합성")

    if "위험성평가" in text or has("위험을 평가", "위험 평가"):
        signals.append("위험성평가")
    if has("작업중지", "작업 중지", "멈춰", "세워", "계속 작업", "계속해") or (
        has("물이", "물은", "물이 새") and has("많아", "늘", "증가")
    ):
        signals.append("작업중지")
    if has("작업 재개", "작업재개", "재개 조건", "재개 승인", "재가동", "다시 시작", "다시 들어", "재입갱", "재운행", "다시 돌"):
        signals.append("작업 재개")
    if has("새 작업", "공정 변경", "작업 방법이 바뀌", "위치 바뀜", "상태가 달라", "조건이 달라", "변경"):
        signals.append("변경 작업")
        signals.append("변경사항")

    if has("의식", "쓰러"):
        signals.append("의식상실")
    if has("응급", "다쳤", "부상"):
        signals.append("응급조치")
    if "구조" in text or "데려오" in text:
        signals.append("구조")
    if has("2차 위험", "바로 들어가", "구조자") or (
        has("밀폐", "갱내") and has("쓰러", "의식") and has("구조", "데려오", "사람부터 빼")
    ):
        signals.append("2차 위험")
    if has("사고", "다쳤", "부상", "쓰러", "칠 뻔", "스쳤"):
        signals.append("사고")

    if has("tbm", "티비엠", "회의 없이"):
        signals.append("TBM")
        signals.append("작업 전 공유")
    if "굴진" in text:
        signals.append("굴진")
    if "천공" in text:
        signals.append("천공")

    if has("사고보고", "사고 보고", "보고해야", "누구에게 보고", "보고 순서", "보고 항목") or (
        has("아차사고", "중대재해", "큰 부상", "중상") and has("해야", "대응", "근거")
    ):
        signals.append("사고보고")
    if "기록" in text or "증거" in text or "자료를" in text or (
        "사고" in text and has("적어", "남겨", "남김")
    ):
        signals.append("기록")
    if has("현장보존", "현장 보존", "현장을 그대로", "현장 치우기 전"):
        signals.append("현장보존")
    if "중대재해" in text or "큰 부상" in text or "중상" in text:
        signals.append("중대재해")
    if "재발방지" in text:
        signals.append("재발방지")
    return _unique(signals)


def _extract_stage(text: str, signals: list[str]) -> str:
    if _contains_any(text, ("교육자료", "교육 자료")):
        return "교육·회의 단계"
    if _contains_any(text, (
        "사고 발생 직후", "사고 직후", "사고가 나", "사고 발생", "사고 원인",
        "쓰러", "의식을 잃", "감전된", "다쳤", "부상 상태", "응급조치 후",
        "칠 뻔", "스쳤", "아차사고", "중대재해", "중상 사고", "큰 부상",
    )):
        return "사고 발생 직후"
    if _contains_any(text, ("정비 전에", "정비전", "정비 시작 전", "점검 전", "점검전", "수리 전", "막힘 제거 전")):
        return "점검·정비 전"
    if {"전원 차단", "잠금·표지", "무전압"}.issubset(set(signals)):
        return "점검·정비 전"
    if _contains_any(text, ("물이", "물은", "물이 새")) and _contains_any(text, ("늘", "많아", "증가")):
        return "이상징후 발견 직후"
    acute_signals = {"출수 증가", "지보 변형", "시야 저하"}
    if acute_signals.intersection(signals) and _contains_any(text, ("늘", "증가", "굵어", "휘", "안 보", "너무 많")):
        return "이상징후 발견 직후"
    if _contains_any(text, (
        "재개 전", "재개하려", "재개 조건", "재개 승인", "다시 들어가기 전",
        "다시 시작", "재입갱", "재운행", "다시 돌리기 전", "다시 들어가",
    )) and "재가동 방지" not in text:
        return "작업 재개 전"
    if _contains_any(text, ("시작 직전", "바로 장약", "투입 직전", "시작해도", "바로 천공", "바로 시작")):
        return "작업 시작 직전"
    if _contains_any(text, ("tbm에서", "티비엠에서", "교육", "회의에서", "공유", "사례를 tbm", "사례와 tbm")) or (
        _contains_any(text, ("사례", "사고가 실제")) and _contains_any(text, ("예방", "교육자료", "지시"))
    ):
        return "교육·회의 단계"
    if _contains_any(text, ("발파 후", "발파 뒤", "발파하고 나서", "발파 끝나고", "작업 후", "작업 끝나고", "정비끝나고")):
        return "작업 후"
    if _contains_any(text, (
        "작업 전", "작업전", "하기 전에", "전에", "전까지", "장약 전", "장약전에",
        "시작하기 전", "작업 방법이 바뀌", "공정 변경", "위험을 평가", "보호구 적합성",
        "후진 작업에서", "고장난 장비", "지급 전", "교체 전",
    )):
        return "작업 전"
    if _contains_any(text, ("tbm", "티비엠")) and "다시 해야" not in text:
        return "교육·회의 단계"
    if _contains_any(text, ("tbm", "티비엠")) and "다시 해야" in text:
        return "작업 전"
    if _contains_any(text, (
        "작업 중", "진행 중", "진행중", "굴진중", "천공 중", "돌고", "도는 중",
        "운행 중", "계속 일", "계속 작업", "계속 천공", "계속해", "옆으로", "안 썼어",
        "청소 잠깐", "정전으로", "멈추고",
    )):
        return "작업 중"
    anomaly_signals = {
        "출수 증가", "지보 변형", "가스 상승", "환기 정지", "산소 부족",
        "누전", "낙반", "시야 저하", "작업중지", "유해가스",
    }
    if anomaly_signals.intersection(signals) or _contains_any(text, (
        "갑자기", "이상", "늘", "증가", "고장", "탄 냄새", "먼지가 난",
    )):
        return "이상징후 발견 직후"
    if "구조 작업" in text and _contains_any(text, ("예방", "공식 근거", "교육")):
        return "교육·회의 단계"
    return "불명확"


def _extract_actor(text: str) -> str:
    actors = (
        ("광산안전관리자", ("광산안전관리자",)),
        ("발파책임자", ("발파책임자", "발파 책임자")),
        ("현장관리자", ("현장관리자", "현장 관리자", "관리자는", "관리자가")),
        ("사업주·경영책임자", ("사업주", "경영책임자")),
        ("구조·응급 대응자", ("구조자", "응급 대응자", "구급")),
        ("작업자", ("작업자", "근로자")),
    )
    for actor, terms in actors:
        if _contains_any(text, terms):
            return actor
    return "불명확"


def _extract_requested_outputs(text: str) -> list[str]:
    requested: list[str] = []
    rules: tuple[tuple[str, tuple[str, ...]], ...] = (
        ("조치 순서", ("순서", "무엇부터", "뭐부터", "먼저 뭘", "먼저 무엇", "어떻게 조치", "지시해야", "지시부터", "차단하고", "차단과", "사람부터", "표지한 다음")),
        ("즉시 조치", ("즉시", "우선조치", "우선 조치", "어떻게 해야", "뭘 해야", "뭐 해야", "먼저 뭐", "먼저 뭘", "사람부터", "바로 들어", "먼저 판단")),
        ("점검항목", ("점검", "확인", "뭘 봐", "무엇을 봐", "재측정", "점검표", "어떤 상태", "고름")),
        ("위험 원인", ("왜", "위험한 이유", "왜 필요한", "이유")),
        ("작업중지 여부", ("작업중지", "작업 중지", "중지 여부", "멈춰", "세워야", "세워야 함", "계속 일", "계속 작업", "계속해", "계속 써")),
        ("대피·출입통제", ("대피", "출입통제", "출입 통제", "사람부터 빼", "사람 빼", "어디로 보내", "접근 통제", "현장 통제", "통제 순서")),
        ("작업 재개 조건", ("재개", "다시 들어", "다시 시작", "재입갱", "재가동", "재운행", "다시 돌", "계속해도", "계속 작업해도")),
        ("보호구", ("보호구", "안전모", "안전대", "마스크", "보호장비")),
        ("보고 대상", ("보고", "신고")),
        ("기록관리", ("기록", "증거", "뭘 남", "무엇을 남", "자료를", "적어야", "남겨")),
        ("공식 법령·지침 근거", ("법적 근거", "공식 근거", "관련 근거", "근거", "법령", "지침")),
        ("공식 사고사례", ("실제 사고", "사고사례", "사고 사례", "사고가 실제", "사례도", "사례 있", "사례를 보여", "사례까지", "사례와")),
        ("유사 위험 사례", ("비슷한 사고", "비슷한 사례", "유사 위험", "유사 사례", "작업중지 사례")),
        ("교육·TBM 공유사항", ("tbm", "티비엠", "교육", "공유", "작업자에게 알")),
        ("사업주·관리자 대응", ("사업주", "관리자", "작업자에게", "지시", "누가", "책임자")),
        ("작업 가능 여부 판단", ("해도 돼", "해도 됨", "면 돼", "가능", "써도 되", "써도 됨", "계속 써", "들어가도", "들어갈지", "가도 돼", "만져도", "지나가도", "시작해도", "운행", "작업시키", "계속해", "계속 일", "계속 천공", "다시 해야")),
        ("예방대책", ("예방", "대처방안", "감소대책", "방지", "막으려")),
    )
    for output, terms in rules:
        if _contains_any(text, terms):
            requested.append(output)
    if not requested:
        requested.append("즉시 조치")
    if "작업 가능 여부 판단" in requested and _contains_any(text, ("계속", "재개", "다시")):
        requested.append("작업 재개 조건")
    if "작업 가능 여부 판단" in requested and _contains_any(text, ("계속", "운행", "써도", "작업시키")):
        requested.append("작업중지 여부")
    if "조치 순서" in requested and _contains_any(text, ("점검", "확인", "위험", "출수", "가스", "설비")):
        requested.append("점검항목")
    if _contains_any(text, ("사례", "사고가 실제")) and "공식 사고사례" not in requested:
        requested.append("공식 사고사례")
    urgent_terms = (
        "불발", "잔류약", "출수", "지보", "감전", "누전", "끼임", "쓰러",
        "의식", "부석", "갑자기", "탄 냄새", "이상음", "안 썼", "사고 원인", "물이 새",
    )
    if _contains_any(text, urgent_terms) and _contains_any(text, ("해야", "어야", "돼", "됨", "먼저", "보내", "보고")):
        requested.append("즉시 조치")
    if _contains_any(text, ("먼저", "시작해도", "계속", "이상")) and _contains_any(text, ("모르", "몰라", "모름", "어떤", "뭘")):
        requested.append("점검항목")
    if _contains_any(text, ("현장보존", "현장 보존", "보고 항목", "아차사고")):
        requested.append("기록관리")
    if "아차사고" in text:
        requested.append("보고 대상")
    if _contains_any(text, ("컨베이어", "벨트", "파쇄기")) and _contains_any(text, ("먼저", "빼", "제거", "세워")):
        requested.append("작업중지 여부")
    if _contains_any(text, ("보호구", "보호장비")) and "교체" in text:
        requested.extend(("사업주·관리자 대응", "작업 가능 여부 판단"))
    return _unique(requested)


def _ambiguity(
    text: str,
    primary: str,
    signals: list[str],
) -> tuple[str, list[str], str]:
    reasons: list[str] = []
    clarification = ""
    generic_water = _contains_any(text, ("물이", "물은", "물이 새", "물이 갑자기")) and not _contains_any(
        text, ("천반", "천장", "측벽", "갱도", "발파공", "발파 구멍", "배관", "누유", "냉각수", "출수", "유출수")
    )
    explicit_unknown = _contains_any(text, ("모르", "몰라", "모름", "불명확", "정확하지", "아직 안 정", "원인은 아직"))
    if generic_water or (_contains_any(text, ("출수", "유출수")) and not _contains_any(text, ("천반", "천장", "측벽", "갱도", "발파공", "배관"))):
        reasons.append("물 발생 위치가 확인되지 않음")
        clarification = "물이 발생하거나 증가한 위치가 천반·측벽·갱도·발파공·배관·장비 중 어디입니까?"
    elif explicit_unknown:
        reasons.append("위험의 원인 또는 대상이 확인되지 않음")
        prompts = {
            DOMAIN_BLASTING: "이상이 확인된 대상이 화약류·뇌관·점화설비 중 무엇이며 젖음과 손상 중 어떤 징후입니까?",
            DOMAIN_VENTILATION: "확인되지 않은 핵심 사항은 가스 종류, 발생 위치, 측정 상태 중 무엇입니까?",
            DOMAIN_ELECTRICAL: "이상 징후가 발생한 전기설비와 현재 차단 상태는 무엇입니까?",
            DOMAIN_CONVEYOR: "이상 징후가 난 설비 위치와 현재 운전·에너지 차단 상태는 무엇입니까?",
            DOMAIN_DUST: "분진의 발생 위치와 현재 집진·살수 작동 상태는 어떻습니까?",
            DOMAIN_PPE: "문제가 된 보호구 종류와 해당 작업의 주요 위험은 무엇입니까?",
            DOMAIN_RISK: "작업을 중지하게 한 이상 징후가 발생한 위치와 대상은 무엇입니까?",
            DOMAIN_EMERGENCY: "부상자의 의식·호흡 상태와 구조 장소의 2차 위험 유무는 어떻습니까?",
            DOMAIN_TBM: "변경되거나 이상이 확인된 작업·장비·작업 위치는 무엇입니까?",
            DOMAIN_REPORTING: "사고의 발생 장소와 인명피해·추가 위험의 현재 상태는 어떻습니까?",
            DOMAIN_MOBILE: "장비의 운전 상태와 작업자 접근 위치는 어디입니까?",
        }
        clarification = prompts.get(primary, "조치가 달라지는 핵심 위험의 발생 위치와 대상은 무엇입니까?")
    elif primary == DOMAIN_MOBILE and len(text) < 36 and _contains_any(text, ("장비", "굴착기", "로더")):
        reasons.append("장비의 운전 상태와 접근 위치가 확인되지 않음")
        clarification = "장비가 운전 중인지와 작업자가 접근하려는 위치는 어디입니까?"
    elif primary == DOMAIN_PPE and "마스크 안" in text and not _contains_any(text, ("분진", "가스", "도장", "용접")):
        reasons.append("보호구가 필요한 작업 위험이 확인되지 않음")
        clarification = "마스크가 필요한 작업의 분진·가스 등 주요 노출 위험은 무엇입니까?"
    elif primary == DOMAIN_TBM and "이상" in text and not _contains_any(text, ("소음", "진동", "누유", "과열", "파손")):
        reasons.append("장비 이상 징후가 구체적이지 않음")
        clarification = "천공 장비에서 확인된 소음·진동·누유·과열 등 구체적인 이상 징후는 무엇입니까?"
    elif primary == DOMAIN_RISK and not set(signals).difference({"작업중지", "작업 재개", "변경 작업", "변경사항"}):
        reasons.append("작업 판단을 바꾸는 구체적인 위험 신호가 확인되지 않음")
        clarification = "작업을 중지하거나 재개하려는 장소에서 확인된 구체적인 이상 징후는 무엇입니까?"
    elif primary == DOMAIN_EMERGENCY and _contains_any(text, ("바로 들어가", "상태가 정확하지", "상태가 정확하지")) and not _contains_any(text, ("가스", "전기", "화재", "낙반")):
        reasons.append("부상 상태 또는 구조 장소의 2차 위험이 확인되지 않음")
        clarification = "부상자의 의식·호흡 상태와 구조 장소의 2차 위험 유무는 어떻습니까?"
    level = "high" if generic_water or (primary == DOMAIN_RISK and explicit_unknown) else ("medium" if reasons else "low")
    return level, reasons, clarification


def _urgency(text: str, signals: list[str], stage: str) -> str:
    high_signals = {
        "불발", "가스 상승", "산소 부족", "출수 증가", "지보 변형", "낙반",
        "감전", "누전", "운전 중 제거", "끼임", "충돌", "의식상실", "중대재해",
    }
    if high_signals.intersection(signals) or stage == "사고 발생 직후":
        return "high"
    if stage in {"작업 중", "이상징후 발견 직후"} or "작업중지" in signals:
        return "elevated"
    return "normal"


def _required_concepts(primary: str, signals: list[str], outputs: list[str], urgency: str) -> list[str]:
    concepts = list(signals)
    pre_blasting_check = primary == DOMAIN_BLASTING and "발파공 침수" in signals and "불발" not in signals
    if urgency == "high" or ("조치 순서" in outputs and not pre_blasting_check):
        concepts.extend(("작업중지", "대피·출입통제", "책임자 보고", "안전한 위치에서 점검"))
    if "작업 재개 조건" in outputs or "작업 가능 여부 판단" in outputs:
        concepts.extend(("위험 제거", "책임자 확인", "작업 재개 판단"))
    if "기록관리" in outputs:
        concepts.append("조치·점검·승인 기록")
    if primary == DOMAIN_ELECTRICAL and _contains_any(" ".join(signals), ("정비", "전원 차단", "잠금·표지", "무전압")):
        concepts.extend(("전원 차단", "잠금·표지", "잔류에너지 제거", "무전압 확인", "재투입 방지"))
    if primary == DOMAIN_VENTILATION and _contains_any(" ".join(signals), ("환기 정지", "가스 상승", "산소 부족")):
        concepts.extend(("환기 복구", "가스·산소 재측정", "재개 승인"))
    if primary == DOMAIN_GROUND and "출수 증가" in signals:
        concepts.extend(("천반·지보·출수·배수 점검", "균열·박리·부석 확인", "배수·집수 검토"))
    if primary == DOMAIN_BLASTING and "발파공 침수" in signals:
        concepts.extend(("기상·낙뢰 확인", "화약류·뇌관 습윤 확인", "배수·접근로 점검", "작업 연기·중지"))
    return _unique(concepts)


def _forbidden_topics(primary: str, signals: list[str], ambiguity: str) -> list[str]:
    forbidden: list[str] = []
    if primary == DOMAIN_BLASTING and "우천" in signals and "불발" not in signals:
        forbidden.append("불발")
    if primary == DOMAIN_BLASTING and "불발" in signals:
        forbidden.append("우천")
    if ambiguity in {"medium", "high"} and _contains_any(" ".join(signals), ("작업중지",)):
        forbidden.extend(("천반으로 확정", "발파공으로 확정"))
    return forbidden


def _display_label(primary: str, signals: list[str], ambiguity: str) -> str:
    base = {
        DOMAIN_BLASTING: "발파/불발",
        DOMAIN_VENTILATION: "환기/유해가스",
        DOMAIN_GROUND: "낙반/붕락/지보",
        DOMAIN_ELECTRICAL: "전기/감전/LOTO",
        DOMAIN_CONVEYOR: "컨베이어/끼임",
        DOMAIN_MOBILE: "차량/장비/충돌",
        DOMAIN_DUST: "분진/호흡보호",
        DOMAIN_PPE: "보호구",
        DOMAIN_RISK: "위험성평가/작업중지",
        DOMAIN_EMERGENCY: "응급/구조",
        DOMAIN_TBM: "TBM/굴진/천공",
        DOMAIN_REPORTING: "사고보고/기록",
        OUT_OF_SCOPE_DOMAIN: "범위 밖",
    }[primary]
    detail = ""
    if "우천" in signals and "발파공 침수" in signals:
        detail = "우천 발파공 침수"
    elif "불발" in signals:
        detail = "발파 후 불발 의심"
    elif "출수 증가" in signals:
        detail = "천반 출수 증가"
    elif "환기 정지" in signals:
        detail = "환기설비 정지"
    elif "운전 중 제거" in signals:
        detail = "운전 중 제거 위험"
    elif ambiguity in {"medium", "high"}:
        detail = "핵심 위치·상태 확인 필요"
    elif signals:
        detail = signals[0]
    return f"{base} · {detail}" if detail else base


def _search_plan(
    original: str,
    primary: str,
    secondary: list[str],
    signals: list[str],
    stage: str,
    outputs: list[str],
    case_requested: bool,
) -> tuple[list[str], str]:
    if primary == OUT_OF_SCOPE_DOMAIN:
        return [original] if original else [], ""
    queries = [original]
    signal_terms = " ".join(signals[:6])
    domain_terms = " ".join([primary, *secondary[:2]])
    queries.append(" ".join(part for part in (domain_terms, signal_terms, stage) if part))
    action_map = {
        "조치 순서": "작업중지 대피 출입통제 보고 점검 개선조치 작업 재개 승인 기록",
        "점검항목": "안전 점검 확인사항",
        "작업 재개 조건": "위험 제거 재측정 책임자 확인 작업 재개 승인",
        "기록관리": "점검 조치 보고 작업 재개 승인 기록",
        "공식 법령·지침 근거": "공식 법령 지침 근거 문서 chunk_id",
        "예방대책": "위험 예방대책 개선조치",
    }
    action_terms = " ".join(action_map[item] for item in outputs if item in action_map)
    if action_terms:
        queries.append(" ".join(part for part in (primary, signal_terms, action_terms) if part))
    case_query = ""
    if case_requested:
        case_query = " ".join(part for part in (primary, signal_terms, "공식 사고사례") if part)
        queries.append(case_query)
    elif "공식 법령·지침 근거" in outputs:
        queries.append(" ".join(part for part in (primary, signal_terms, "공식 법령 지침") if part))
    return _unique(queries)[:4], case_query


def analyze_query(query: str) -> QueryUnderstanding:
    original = str(query or "")
    normalized = normalize_query(original)
    text = normalized.lower()
    primary, secondary, scores = _classify_domains(text)
    in_scope = primary != OUT_OF_SCOPE_DOMAIN
    signals = _extract_signals(text) if in_scope else []
    stage = _extract_stage(text, signals) if in_scope else "불명확"
    actor = _extract_actor(text) if in_scope else "불명확"
    outputs = _extract_requested_outputs(text) if in_scope else []
    ambiguity_level, ambiguity_reasons, clarification = (
        _ambiguity(text, primary, signals) if in_scope else ("low", [], "")
    )
    urgency = _urgency(text, signals, stage) if in_scope else "normal"
    case_requested = bool({"공식 사고사례", "유사 위험 사례"}.intersection(outputs))
    required = _required_concepts(primary, signals, outputs, urgency) if in_scope else []
    forbidden = _forbidden_topics(primary, signals, ambiguity_level) if in_scope else []
    search_queries, case_query = _search_plan(
        normalized,
        primary,
        secondary,
        signals,
        stage,
        outputs,
        case_requested,
    )
    return QueryUnderstanding(
        original_query=original,
        normalized_query=normalized,
        primary_domain=primary,
        secondary_domains=secondary,
        hazard_signals=signals,
        work_stage=stage,
        actor=actor,
        requested_outputs=outputs,
        urgency=urgency,
        ambiguity_level=ambiguity_level,
        ambiguity_reasons=ambiguity_reasons,
        clarification_question=clarification,
        search_queries=search_queries,
        case_search_query=case_query,
        required_concepts=required,
        forbidden_dominant_topics=forbidden,
        display_label=_display_label(primary, signals, ambiguity_level),
        official_case_requested=case_requested,
        in_scope=in_scope,
    )


def rerank_results_by_understanding(
    understanding: QueryUnderstanding,
    candidates: list[dict[str, Any]],
    top_k: int = 5,
) -> list[dict[str, Any]]:
    """원시 distance를 보존하고 질문 문맥 신호로만 제한적으로 재정렬합니다."""
    aliases = {
        "발파공 침수": ("발파공", "침수", "공내수"),
        "화약류 습윤": ("화약", "폭약", "뇌관", "습윤", "젖"),
        "환기 정지": ("환기", "송풍", "정지"),
        "가스 상승": ("가스", "측정", "농도"),
        "출수 증가": ("출수", "유출수", "지하수"),
        "지보 변형": ("지보", "변형", "손상"),
        "잠금·표지": ("잠금", "표지", "loto"),
        "운전 중 제거": ("운전", "가동", "제거", "청소"),
        "작업 재개": ("재개", "재가동", "승인"),
    }
    requested_terms = {
        "조치 순서": ("작업중지", "대피", "통제", "보고", "점검"),
        "점검항목": ("점검", "확인"),
        "작업 재개 조건": ("재개", "승인", "재측정"),
        "기록관리": ("기록", "점검표"),
        "공식 법령·지침 근거": ("법", "지침", "기준"),
    }
    stage_terms = {
        "작업 전": ("작업 전", "사전", "점검"),
        "작업 시작 직전": ("작업 전", "시작 전", "점검"),
        "작업 중": ("작업 중", "운전 중", "가동 중"),
        "이상징후 발견 직후": ("이상", "징후", "즉시", "작업중지"),
        "작업 후": ("작업 후", "발파 후", "종료 후"),
        "사고 발생 직후": ("사고", "응급", "보고"),
        "점검·정비 전": ("정비 전", "점검 전", "차단"),
        "작업 재개 전": ("재개", "재가동", "승인"),
        "교육·회의 단계": ("교육", "tbm", "회의"),
    }
    seen: set[str] = set()
    scored: list[tuple[float, int, dict[str, Any]]] = []
    for index, candidate in enumerate(candidates):
        chunk_id = str(candidate.get("chunk_id", "")).strip()
        source = str(candidate.get("source", "")).strip()
        text = " ".join((source, str(candidate.get("text", "")))).lower()
        key = chunk_id or f"{source}|{text[:120]}"
        if key in seen:
            continue
        seen.add(key)
        distance = candidate.get("distance")
        base = float(distance) if isinstance(distance, (int, float)) else 99.0
        context_bonus = 0.0
        signal_hits = 0
        for signal in understanding.hazard_signals:
            terms = aliases.get(signal, (signal.lower(),))
            if any(term in text for term in terms):
                signal_hits += 1
        context_bonus += min(signal_hits, 5) * 0.07
        if any(term in text for term in stage_terms.get(understanding.work_stage, ())):
            context_bonus += 0.04
        request_hits = sum(
            any(term in text for term in requested_terms.get(output, ()))
            for output in understanding.requested_outputs
        )
        context_bonus += min(request_hits, 3) * 0.03
        metadata = candidate.get("metadata")
        if source and chunk_id and isinstance(metadata, dict):
            context_bonus += 0.02
        rerank_value = base - min(context_bonus, 0.45)
        item = dict(candidate)
        item["context_rerank_value"] = rerank_value
        item["context_signal_hits"] = signal_hits
        item["distance"] = distance
        vector_rank = int(item.get("vector_rank") or index + 1)
        scored.append((rerank_value, vector_rank, item))
    scored.sort(key=lambda row: (row[0], row[1]))
    selected = [item for _, _, item in scored[: max(0, top_k)]]
    for rank, item in enumerate(selected, start=1):
        item["rank"] = rank
    return selected


__all__ = [
    "QueryUnderstanding",
    "SUPPORTED_DOMAINS",
    "OUT_OF_SCOPE_DOMAIN",
    "analyze_query",
    "normalize_query",
    "rerank_results_by_understanding",
]
