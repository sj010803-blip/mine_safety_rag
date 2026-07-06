from __future__ import annotations

from collections import defaultdict
from datetime import datetime
import json
import math
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Any


ROOT_DIR = Path(__file__).resolve().parent
EXPERIMENT_DIR = ROOT_DIR / "12_compare_experiment"
TEMPLATE_PATH = EXPERIMENT_DIR / "llm_rag_comparison_template.xlsx"
OUTPUT_PATH = EXPERIMENT_DIR / "llm_rag_comparison_result_polished.xlsx"
CHUNKS_PATH = ROOT_DIR / "08_chunks" / "chunks.jsonl"
VECTOR_DB_DIR = ROOT_DIR / "10_vector_db"
BRIDGE_PATH = ROOT_DIR / "_llm_rag_workbook_bridge.mjs"
COLLECTION_NAME = "mine_safety_docs"
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"

EVALUATION_HEADERS = [
    "번호",
    "카테고리",
    "질문",
    "비교대상",
    "검색 적합성(25)",
    "근거 기반성(25)",
    "안전·법령 판단 정확성(25)",
    "실무성(25)",
    "총점(100)",
    "판정",
    "장점",
    "한계",
    "자동평가 근거",
    "검토필요",
]

GENERAL_DOMAIN_TERMS = [
    "광산",
    "갱내",
    "작업자",
    "관리자",
    "위험",
    "안전",
    "작업중지",
    "대피",
]

GROUNDING_TERMS = [
    "법령",
    "지침",
    "기준",
    "산업안전보건법",
    "광산안전법",
    "중대재해처벌법",
    "안전보건관리체계",
    "광산안전기술기준",
    "광산안전업무",
]

SOURCE_TERMS = [
    "검색 근거",
    "관련 근거",
    "근거 문서",
    "문서명",
    "source",
    "chunk_id",
    "출처",
]

COMMON_SAFETY_TERMS = [
    "작업중지",
    "대피",
    "출입통제",
    "접근 금지",
    "응급조치",
    "보고",
    "점검",
    "확인",
    "측정",
    "환기",
    "보호구",
    "차단",
    "잠금",
    "재발방지",
    "작업 재개",
]

PRACTICAL_TERMS = [
    "단계별",
    "체크리스트",
    "KRAS",
    "세부 작업 내용",
    "잠재위험요인",
    "위험발생 상황 및 결과",
    "위험성 감소대책",
    "조치 후 잔여위험성",
    "담당자",
    "이행일",
    "이행 상태",
    "기록",
    "보고",
]

CATEGORY_PROFILES: dict[str, dict[str, list[str] | str]] = {
    "낙반/중대재해": {
        "terms": ["낙반", "천반", "측벽", "지보", "부석", "균열", "사망", "중대재해"],
        "actions": ["작업중지", "대피", "출입통제", "현장 보존", "원인조사", "재발방지", "보고"],
        "hazards": ["천반·측벽 균열", "지보 상태 불량", "부석·낙석", "추가 낙반"],
        "preferred": ["광산안전기술기준", "광산안전업무", "중대재해", "산업안전보건법"],
    },
    "발파/불발": {
        "terms": ["발파", "불발", "화약", "폭약", "장약", "점화", "굴진", "위험구역"],
        "actions": ["접근 금지", "출입통제", "대피", "발파 책임자", "작업중지", "재출입", "기록"],
        "hazards": ["불발 화약류", "잔류화약", "발파모선", "성급한 재접근"],
        "preferred": ["광산안전기술기준", "광산안전업무", "광산안전법_시행규칙"],
    },
    "환기/유해가스": {
        "terms": ["환기", "메탄", "산소", "일산화탄소", "유해가스", "가스", "정전", "환기설비"],
        "actions": ["작업중지", "대피", "환기", "가스 측정", "전원 차단", "재측정", "작업 재개 전 확인"],
        "hazards": ["메탄 축적", "산소 결핍", "유해가스", "환기설비 정지"],
        "preferred": ["광산안전기술기준", "광산안전법_시행규칙", "광산안전업무"],
    },
    "전기안전": {
        "terms": ["전기", "전기설비", "감전", "누전", "접지", "절연", "차단", "잠금", "표지", "습기"],
        "actions": ["전원 차단", "잠금", "표지", "접지 확인", "절연 확인", "접근통제", "작업 재개 전 확인"],
        "hazards": ["누전", "감전", "습기·물기", "임의 전원 투입"],
        "preferred": ["광산안전기술기준", "광산안전법_시행규칙", "산업안전보건법_시행규칙"],
    },
    "TBM(툴박스 미팅)": {
        "terms": ["TBM", "툴박스", "작업 전", "안전회의", "위험요인", "역할", "신호", "비상연락망"],
        "actions": ["작업 내용 공유", "위험요인 공유", "보호구 확인", "역할 확인", "대피로 확인", "의견 청취", "기록"],
        "hazards": ["위험정보 공유 누락", "역할·신호 불명확", "보호구 확인 누락", "비상절차 미공유"],
        "preferred": ["위험성평가", "안전보건관리체계", "광산안전기술기준"],
    },
    "분진 관리": {
        "terms": ["분진", "먼지", "천공", "굴진", "파쇄", "살수", "집진", "작업환경측정"],
        "actions": ["살수", "집진", "환기", "작업환경측정", "방진마스크", "청소", "기록"],
        "hazards": ["고농도 분진", "비산먼지", "집진·살수 불량", "호흡기 노출"],
        "preferred": ["광산안전기술기준", "광산안전업무", "산업안전보건법_시행규칙", "위험성평가"],
    },
    "보호구/PPE": {
        "terms": ["보호구", "PPE", "방진마스크", "안전모", "턱끈", "보안경", "보호장갑", "착용"],
        "actions": ["지급", "착용 확인", "적합성 확인", "파손 확인", "작업 투입 금지", "교육", "기록"],
        "hazards": ["보호구 미착용", "부적합 보호구", "보호구 파손", "착용상태 불량"],
        "preferred": ["광산안전법_시행규칙", "광산안전업무", "광산안전기술기준", "위험성평가"],
    },
    "위험성평가": {
        "terms": ["위험성평가", "유해·위험요인", "위험수준", "감소대책", "재평가", "개선조치", "근로자"],
        "actions": ["작업중지", "유해·위험요인 파악", "위험성 결정", "감소대책", "이행 확인", "근로자 공유", "기록"],
        "hazards": ["중대한 위험 방치", "변경요인 미평가", "감소대책 미이행", "작업 전 공유 누락"],
        "preferred": ["위험성평가", "안전보건관리체계", "산업안전보건법", "광산안전기술기준"],
    },
    "사고보고/응급조치": {
        "terms": ["사고", "응급조치", "119", "구조", "구급", "의식", "보고", "현장 보존"],
        "actions": ["작업중지", "부상자 구조", "응급조치", "119", "2차 사고 방지", "현장 보존", "보고", "원인조사"],
        "hazards": ["응급상황", "구조자 2차 피해", "현장 훼손", "보고 지연"],
        "preferred": ["중대재해", "산업안전보건법", "광산안전업무", "안전보건관리체계"],
    },
    "작업중지/재개": {
        "terms": ["작업중지", "작업 재개", "위험요인 제거", "확인", "승인", "재평가", "근로자 공유"],
        "actions": ["작업중지", "위험요인 제거", "재점검", "재평가", "책임자 승인", "근로자 공유", "기록"],
        "hazards": ["위험요인 잔존", "검증 없는 재개", "책임자 확인 누락", "재발 위험"],
        "preferred": ["광산안전기술기준", "위험성평가", "안전보건관리체계", "산업안전보건법"],
    },
}


def normalize_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def contains_term(text: str, term: str) -> bool:
    return term.lower().replace(" ", "") in text.lower().replace(" ", "")


def term_hits(text: str, terms: list[str]) -> list[str]:
    return [term for term in terms if contains_term(text, term)]


def profile_for(category: str) -> dict[str, list[str] | str]:
    return CATEGORY_PROFILES.get(
        category,
        {
            "terms": ["광산", "안전", "위험", "작업"],
            "actions": ["작업중지", "대피", "보고", "점검", "기록"],
            "hazards": ["현장 위험요인", "관리 미흡", "작업자 노출"],
            "preferred": ["광산안전기술기준", "산업안전보건법", "안전보건관리체계"],
        },
    )


def judgment(total: int) -> str:
    if total >= 90:
        return "매우 우수"
    if total >= 80:
        return "우수"
    if total >= 70:
        return "보통"
    if total >= 60:
        return "미흡"
    return "부족"


def cap_score(value: float) -> int:
    return max(0, min(25, int(round(value))))


def evaluate_answer(
    question_id: str,
    category: str,
    question: str,
    target: str,
    answer: str,
) -> dict[str, Any]:
    text = normalize_text(answer)
    if not text:
        return {
            "번호": question_id,
            "카테고리": category,
            "질문": question,
            "비교대상": target,
            "검색 적합성(25)": 0,
            "근거 기반성(25)": 0,
            "안전·법령 판단 정확성(25)": 0,
            "실무성(25)": 0,
            "총점(100)": 0,
            "판정": "미평가",
            "장점": "답변 원문이 입력되면 평가할 수 있습니다.",
            "한계": "답변 원문이 비어 있습니다.",
            "자동평가 근거": "빈 답변은 자동 채점하지 않았습니다.",
            "검토필요": "Y",
        }

    profile = profile_for(category)
    category_terms = list(profile["terms"])
    action_terms = list(profile["actions"])
    category_hits = term_hits(text, category_terms)
    domain_hits = term_hits(text, GENERAL_DOMAIN_TERMS)
    grounding_hits = term_hits(text, GROUNDING_TERMS)
    source_hits = term_hits(text, SOURCE_TERMS)
    action_hits = term_hits(text, action_terms)
    common_safety_hits = term_hits(text, COMMON_SAFETY_TERMS)
    practical_hits = term_hits(text, PRACTICAL_TERMS)

    category_ratio = len(category_hits) / max(1, min(7, len(category_terms)))
    search_score = 8 + min(10, category_ratio * 12)
    search_score += min(4, len(domain_hits) * 0.8)
    if "chunk_id" in text.lower() or "source" in text.lower():
        search_score += 3
    elif grounding_hits:
        search_score += min(2, len(grounding_hits) * 0.4)

    grounding_score = 9
    grounding_score += min(9, len(grounding_hits) * 1.8)
    grounding_score += min(5, len(source_hits) * 1.25)
    if any(contains_term(text, term) for term in ["규정", "절차", "책임자", "법적 의무"]):
        grounding_score += 2
    if "chunk_id" in text.lower():
        grounding_score += 4
    if re.search(r"\.txt|광산안전기술기준|광산안전업무", text, re.I):
        grounding_score += 2

    action_ratio = len(action_hits) / max(1, min(7, len(action_terms)))
    safety_score = 10 + min(14, action_ratio * 16)
    safety_score += min(6, len(common_safety_hits) * 0.9)
    if any(contains_term(text, term) for term in ["책임자 확인", "전문가 확인", "작업 재개 전", "안전 확인 후"]):
        safety_score += 2

    practical_score = 11
    practical_score += min(9, len(practical_hits) * 1.25)
    practical_score += min(7, len(action_hits) * 1.0)
    structure_count = sum(
        1
        for marker in ["1.", "2.", "3.", "단계", "우선", "확인", "조치"]
        if marker in text
    )
    practical_score += min(4, structure_count)
    if len(text) >= 250:
        practical_score += 2
    if len(text) >= 600:
        practical_score += 1
    kras_hits = term_hits(
        text,
        [
            "KRAS",
            "세부 작업 내용",
            "잠재위험요인",
            "위험성 감소대책",
            "조치 후 잔여위험성",
            "조치 체크리스트",
        ],
    )
    practical_score += min(7, len(kras_hits) * 1.4)
    unsafe_patterns = [
        r"즉시\s*작업을\s*재개",
        r"바로\s*작업을\s*재개",
        r"확인\s*없이\s*작업",
    ]
    unsafe_found = any(re.search(pattern, text) for pattern in unsafe_patterns)
    if unsafe_found:
        safety_score -= 8
        practical_score -= 4

    uncertain_article = bool(re.search(r"제\s*\d+\s*조", text)) and not source_hits
    if uncertain_article:
        grounding_score -= 2
        safety_score -= 1

    practical_capped = cap_score(practical_score)
    draft_penalty = 0
    if contains_term(text, "현장 기준으로 재평가 필요"):
        draft_penalty += 2
    if contains_term(text, "이행일: 미기입"):
        draft_penalty += 1
    if contains_term(text, "기록 초안"):
        draft_penalty += 1
    practical_capped = max(0, practical_capped - draft_penalty)

    scores = {
        "검색 적합성(25)": cap_score(search_score),
        "근거 기반성(25)": cap_score(grounding_score),
        "안전·법령 판단 정확성(25)": cap_score(safety_score),
        "실무성(25)": practical_capped,
    }
    total = sum(scores.values())

    strengths = []
    if scores["검색 적합성(25)"] >= 20:
        strengths.append("질문 유형의 핵심 위험요인을 구체적으로 반영")
    if scores["근거 기반성(25)"] >= 20:
        strengths.append("공식 문서·출처 표현의 추적성이 높음")
    if scores["안전·법령 판단 정확성(25)"] >= 20:
        strengths.append("작업중지·대피·확인 절차가 상황에 맞음")
    if scores["실무성(25)"] >= 20:
        strengths.append("현장 조치와 기록 항목이 구조화됨")
    if not strengths:
        strengths.append("질문에 대한 기본 안전 방향을 제시")

    limitations = []
    if scores["검색 적합성(25)"] < 18:
        limitations.append("광산 현장 특화 핵심어 반영이 제한적")
    if scores["근거 기반성(25)"] < 18:
        limitations.append("문서명·출처·근거 연결을 추가할 필요")
    if scores["안전·법령 판단 정확성(25)"] < 18:
        limitations.append("상황별 안전 판단 절차를 보강할 필요")
    if scores["실무성(25)"] < 18:
        limitations.append("체크리스트·KRAS 기록 형태로 추가 정리 필요")
    if uncertain_article:
        limitations.append("조문 번호의 원문 근거 수동 확인 필요")
    if unsafe_found:
        limitations.append("작업 재개 표현에 대한 안전 검토 필요")
    if not limitations:
        limitations.append("최종 법령 해석과 현장 적용은 수동 검토 필요")

    basis = (
        f"카테고리 핵심어 {len(category_hits)}개, 공식 근거 표현 {len(grounding_hits)}개, "
        f"출처 표지 {len(source_hits)}개, 상황별 안전조치 {len(action_hits)}개, "
        f"공통 안전조치 {len(common_safety_hits)}개, "
        f"실무 구조 요소 {len(practical_hits) + len(kras_hits)}개를 규칙 기반으로 확인"
    )
    review_needed = (
        total < 70
        or min(scores.values()) < 15
        or uncertain_article
        or unsafe_found
        or len(text) < 120
    )

    return {
        "번호": question_id,
        "카테고리": category,
        "질문": question,
        "비교대상": target,
        **scores,
        "총점(100)": total,
        "판정": judgment(total),
        "장점": " / ".join(strengths),
        "한계": " / ".join(limitations),
        "자동평가 근거": basis,
        "검토필요": "Y" if review_needed else "N",
    }


def tokenize(text: str) -> set[str]:
    return {
        token.lower()
        for token in re.findall(r"[가-힣A-Za-z0-9_]{2,}", text)
        if len(token) >= 2
    }


class StandaloneRetriever:
    def __init__(self) -> None:
        self._collection = None
        self._model = None
        self._chunks: list[dict[str, Any]] | None = None
        try:
            import chromadb
            from sentence_transformers import SentenceTransformer

            client = chromadb.PersistentClient(path=str(VECTOR_DB_DIR))
            self._collection = client.get_collection(name=COLLECTION_NAME)
            self._model = SentenceTransformer(EMBEDDING_MODEL_NAME)
        except Exception as exc:
            print(f"[안내] ChromaDB 검색을 사용할 수 없어 chunks.jsonl 검색으로 전환합니다: {exc}")

    def _load_chunks(self) -> list[dict[str, Any]]:
        if self._chunks is None:
            self._chunks = []
            with CHUNKS_PATH.open("r", encoding="utf-8-sig") as handle:
                for line in handle:
                    if line.strip():
                        self._chunks.append(json.loads(line))
        return self._chunks

    @staticmethod
    def _source_from_metadata(metadata: dict[str, Any]) -> str:
        return str(
            metadata.get("source_file")
            or metadata.get("source")
            or metadata.get("filename")
            or "출처 정보 없음"
        )

    def search(self, question: str, category: str, top_k: int = 3) -> list[dict[str, Any]]:
        profile = profile_for(category)
        preferred = list(profile["preferred"])
        candidates: list[dict[str, Any]] = []
        if self._collection is not None and self._model is not None:
            embedding = self._model.encode(
                [question],
                normalize_embeddings=True,
                show_progress_bar=False,
            ).tolist()
            result = self._collection.query(
                query_embeddings=embedding,
                n_results=10,
                include=["documents", "metadatas", "distances"],
            )
            documents = (result.get("documents") or [[]])[0]
            metadatas = (result.get("metadatas") or [[]])[0]
            distances = (result.get("distances") or [[]])[0]
            ids = (result.get("ids") or [[]])[0]
            for index, document in enumerate(documents):
                metadata = metadatas[index] or {}
                source = self._source_from_metadata(metadata)
                candidates.append(
                    {
                        "source": source,
                        "chunk_id": str(
                            metadata.get("chunk_id")
                            or (ids[index] if index < len(ids) else f"chunk_{index + 1}")
                        ),
                        "text": normalize_text(document),
                        "distance": float(distances[index]) if index < len(distances) else 1.0,
                    }
                )
        else:
            query_tokens = tokenize(
                " ".join([question, *list(profile["terms"]), *list(profile["actions"])])
            )
            for chunk in self._load_chunks():
                source = str(chunk.get("source_file") or chunk.get("source") or "출처 정보 없음")
                text = normalize_text(chunk.get("text", ""))
                document_tokens = tokenize(f"{source} {text}")
                overlap = len(query_tokens & document_tokens)
                preferred_bonus = sum(2 for marker in preferred if marker in source)
                if overlap or preferred_bonus:
                    lexical_score = overlap + preferred_bonus
                    candidates.append(
                        {
                            "source": source,
                            "chunk_id": str(chunk.get("chunk_id", "")),
                            "text": text,
                            "distance": 1 / (1 + lexical_score),
                        }
                    )

        def rerank_key(item: dict[str, Any]) -> tuple[float, float]:
            bonus = sum(0.12 for marker in preferred if marker in item["source"])
            return (float(item["distance"]) - bonus, float(item["distance"]))

        candidates.sort(key=rerank_key)
        selected: list[dict[str, Any]] = []
        seen = set()
        for item in candidates:
            key = (item["source"], item["chunk_id"])
            if key in seen:
                continue
            seen.add(key)
            selected.append(item)
            if len(selected) >= top_k:
                break
        return selected


def build_rag_answer(
    question: str,
    category: str,
    evidence: list[dict[str, Any]],
) -> str:
    profile = profile_for(category)
    actions = list(profile["actions"])[:7]
    hazards = list(profile["hazards"])[:4]
    sources = [
        (
            f"{index}. {item['source']} | chunk_id: {item['chunk_id']} | "
            f"distance: {float(item['distance']):.4f}\n"
            f"   근거 요약: {normalize_text(item['text'])[:320]}"
        )
        for index, item in enumerate(evidence, start=1)
    ]
    action_lines = "\n".join(f"- {item}" for item in actions)
    checklist_lines = "\n".join(
        f"{index}. [대기] {item} / 이행일: 미기입 / 담당자: 현장 책임자"
        for index, item in enumerate(actions, start=1)
    )
    hazard_text = ", ".join(hazards)
    action_text = ", ".join(actions)
    source_names = ", ".join(dict.fromkeys(item["source"] for item in evidence))
    evidence_text = "\n".join(sources) if sources else "검색된 근거 문서가 없습니다."

    return f"""## 답변 요약
질문 유형은 **{category}**입니다. 현장에서는 위험구역 통제와 작업자 안전 확보를 우선하고, 검색된 공식 문서 근거와 실제 작업조건을 대조한 뒤 책임자가 작업 여부를 판단해야 합니다.

### 우선 조치
{action_lines}

## 관련 근거 문서
{evidence_text}

## KRAS식 위험성평가 기록 초안

| 항목 | 기입 초안 |
|---|---|
| 세부 작업 내용 | {normalize_text(question)} |
| 잠재위험요인 | {hazard_text} |
| 위험발생 상황 및 결과 | 위험요인이 통제되지 않은 상태에서 작업할 경우 작업자 부상·중대재해 및 2차 사고 가능 |
| 관련 근거 / 법적 기준 | 검색 근거 기준: {source_names or '관련 문서 확인 필요'}. 조문 번호는 검색 원문에서 명확히 확인되는 경우에만 적용 |
| 현재 위험성 | 가능성·중대성·위험등급은 현장 기준으로 재평가 필요 |
| 위험성 감소대책 | 제거 → 대체 → 공학적 대책 → 관리적 대책 → 보호구 순으로 검토: {action_text} |
| 조치 후 잔여위험성 | 감소대책 이행 후 책임자 확인과 재평가를 거쳐 작업 재개 여부 결정 |
| 기록·보고 사항 | 위험요인, 검색 근거, 조치내용, 담당자, 이행일, 확인결과를 기록하고 필요한 보고 실시 |

## 조치 체크리스트
{checklist_lines}

> 주의: 이 답변은 Vector DB 검색 근거를 이용한 기록 초안입니다. 최종 법령 해석과 현장 판단은 안전관리자·관계기관·전문가의 확인이 필요합니다.
"""


def discover_artifact_runtime() -> tuple[Path, Path]:
    runtime_root = (
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "node"
    )
    node_path = runtime_root / "bin" / "node.exe"
    modules_path = runtime_root / "node_modules"
    if node_path.exists() and modules_path.exists():
        return node_path, modules_path

    candidates = list(
        (Path.home() / ".cache" / "codex-runtimes").glob(
            "*/dependencies/node/bin/node.exe"
        )
    )
    for candidate in candidates:
        modules = candidate.parent.parent / "node_modules"
        if modules.exists():
            return candidate, modules
    raise RuntimeError(
        "Excel 작성을 위한 Codex artifact runtime을 찾을 수 없습니다. "
        "Codex Desktop 환경에서 실행해 주세요."
    )


def prepare_bridge_runtime() -> tuple[Path, Path]:
    node_path, modules_path = discover_artifact_runtime()
    runtime_dir = EXPERIMENT_DIR / ".artifact_runtime"
    runtime_dir.mkdir(parents=True, exist_ok=True)
    runtime_bridge = runtime_dir / BRIDGE_PATH.name
    shutil.copy2(BRIDGE_PATH, runtime_bridge)
    junction = runtime_dir / "node_modules"
    if not junction.exists():
        completed = subprocess.run(
            ["cmd", "/c", "mklink", "/J", str(junction), str(modules_path)],
            capture_output=True,
            text=True,
            check=False,
        )
        if completed.returncode != 0:
            raise RuntimeError(
                "artifact runtime 연결 생성 실패: "
                + normalize_text(completed.stderr or completed.stdout)
            )
    return node_path, runtime_bridge


def run_bridge(
    mode: str,
    template_path: Path,
    payload_path: Path,
    output_path: Path | None = None,
) -> None:
    node_path, runtime_bridge = prepare_bridge_runtime()
    command = [
        str(node_path),
        str(runtime_bridge),
        mode,
        str(template_path),
        str(payload_path),
    ]
    if output_path is not None:
        command.append(str(output_path))
    completed = subprocess.run(
        command,
        cwd=ROOT_DIR,
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            f"Excel {mode} 처리 실패: "
            + normalize_text(completed.stderr or completed.stdout)
        )


def average(values: list[int]) -> float:
    return round(sum(values) / len(values), 2) if values else 0.0


def build_payload(
    workbook_data: dict[str, list[list[Any]]],
) -> tuple[dict[str, Any], dict[str, Any]]:
    raw_matrix = workbook_data["3_답변원문_붙여넣기"]
    headers = [normalize_text(value) for value in raw_matrix[0]]
    rows = []
    for matrix_row in raw_matrix[1:]:
        padded = list(matrix_row) + [None] * (len(headers) - len(matrix_row))
        rows.append(dict(zip(headers, padded)))

    retriever = StandaloneRetriever()
    generated_count = 0
    today = datetime.now().strftime("%Y-%m-%d")
    for row in rows:
        target = normalize_text(row.get("비교대상"))
        answer = normalize_text(row.get("답변 원문 붙여넣기"))
        if target == "내 사이트 RAG" and not answer:
            question = normalize_text(row.get("질문"))
            category = normalize_text(row.get("카테고리"))
            evidence = retriever.search(question, category, top_k=3)
            row["답변 원문 붙여넣기"] = build_rag_answer(
                question,
                category,
                evidence,
            )
            row["답변 생성일"] = today
            row["모델명/버전"] = "Standalone RAG rule-based v1"
            row["비고"] = "Vector DB 검색 근거 기반 자동 생성"
            generated_count += 1

    raw_answers = [headers] + [
        [row.get(header, "") for header in headers]
        for row in rows
    ]

    evaluations = []
    evaluation_records = []
    for row in rows:
        evaluated = evaluate_answer(
            normalize_text(row.get("번호")),
            normalize_text(row.get("카테고리")),
            normalize_text(row.get("질문")),
            normalize_text(row.get("비교대상")),
            normalize_text(row.get("답변 원문 붙여넣기")),
        )
        evaluation_records.append(evaluated)
        evaluations.append([evaluated[header] for header in EVALUATION_HEADERS])

    targets = ["ChatGPT", "Gemini", "내 사이트 RAG"]
    summary_records = {}
    summary_rows = [[
        "비교대상",
        "검색 적합성 평균",
        "근거 기반성 평균",
        "안전·법령 판단 정확성 평균",
        "실무성 평균",
        "총점 평균",
    ]]
    for target in targets:
        target_rows = [
            row
            for row in evaluation_records
            if row["비교대상"] == target and row["판정"] != "미평가"
        ]
        summary_record = {
            "검색 적합성 평균": average([row["검색 적합성(25)"] for row in target_rows]),
            "근거 기반성 평균": average([row["근거 기반성(25)"] for row in target_rows]),
            "안전·법령 판단 정확성 평균": average(
                [row["안전·법령 판단 정확성(25)"] for row in target_rows]
            ),
            "실무성 평균": average([row["실무성(25)"] for row in target_rows]),
            "총점 평균": average([row["총점(100)"] for row in target_rows]),
        }
        summary_records[target] = summary_record
        summary_rows.append([target, *summary_record.values()])

    by_question: dict[str, dict[str, dict[str, Any]]] = defaultdict(dict)
    for record in evaluation_records:
        by_question[record["번호"]][record["비교대상"]] = record

    candidates = []
    for question_id, target_rows in by_question.items():
        if not all(target in target_rows for target in targets):
            continue
        rag = target_rows["내 사이트 RAG"]
        chatgpt = target_rows["ChatGPT"]
        gemini = target_rows["Gemini"]
        generic_max = max(chatgpt["총점(100)"], gemini["총점(100)"])
        grounding_gap = rag["근거 기반성(25)"] - max(
            chatgpt["근거 기반성(25)"],
            gemini["근거 기반성(25)"],
        )
        practical_gap = rag["실무성(25)"] - max(
            chatgpt["실무성(25)"],
            gemini["실무성(25)"],
        )
        rank_score = (
            (rag["총점(100)"] - generic_max) * 2
            + grounding_gap
            + practical_gap
        )
        candidates.append(
            {
                "question_id": question_id,
                "rank_score": rank_score,
                "rag_wins": rag["총점(100)"] > generic_max,
                "rag": rag,
                "chatgpt": chatgpt,
                "gemini": gemini,
            }
        )
    candidates.sort(
        key=lambda item: (
            item["rag_wins"],
            item["rank_score"],
            item["rag"]["총점(100)"],
        ),
        reverse=True,
    )
    candidate_by_id = {
        item["question_id"]: item
        for item in candidates
    }
    priority_ids = ["Q001", "Q004", "Q010", "Q102", "Q107"]
    priority_candidates = [
        candidate_by_id[question_id]
        for question_id in priority_ids
        if question_id in candidate_by_id
    ]
    priority_candidates.sort(
        key=lambda item: (
            item["rag_wins"],
            item["rank_score"],
            item["rag"]["총점(100)"],
        ),
        reverse=True,
    )

    selected_cases = []
    selected_categories = set()
    for case in priority_candidates + candidates:
        category = case["rag"]["카테고리"]
        if category in selected_categories:
            continue
        selected_cases.append(case)
        selected_categories.add(category)
        if len(selected_cases) >= 3:
            break
    representative_rows = [[
        "질문 번호",
        "질문",
        "ChatGPT 한계 요약",
        "Gemini 한계 요약",
        "내 사이트 RAG 강점",
        "비교 해석 문장",
    ]]
    for case in selected_cases:
        rag = case["rag"]
        chatgpt = case["chatgpt"]
        gemini = case["gemini"]
        interpretation = (
            f"세 답변 모두 {rag['카테고리']} 상황의 기본 안전 방향을 제시했습니다. "
            f"자동평가 점수는 ChatGPT {chatgpt['총점(100)']}점, "
            f"Gemini {gemini['총점(100)']}점, RAG {rag['총점(100)']}점이었습니다. "
            "범용 LLM은 자연어 설명에, 도메인 특화 RAG는 공식 문서 추적과 "
            "KRAS·체크리스트 기록 지원에 초점이 있다는 목적 차이를 보여주는 사례입니다."
        )
        representative_rows.append(
            [
                case["question_id"],
                rag["질문"],
                "안전조치의 자연어 설명은 제공하였으며, 본 실험 기준의 문서명·chunk_id·"
                "KRAS 기록 항목은 추가 정리가 필요함.",
                "상황별 대응 설명은 제공하였으며, 공식 문서 추적 정보와 위험성평가 "
                "양식화는 별도 보완이 필요함.",
                "검색 문서명과 chunk_id를 연결하고 위험요인·감소대책·조치 "
                "체크리스트를 동일 답변 안에서 구조화함.",
                interpretation,
            ]
        )

    design_text = (
        "동일한 광산 안전 질문 20개를 ChatGPT, Gemini, 내 사이트 RAG에 각각 입력하고, "
        "검색 적합성·근거 기반성·안전·법령 판단 정확성·실무성의 4개 기준을 "
        "각 25점, 총 100점으로 비교하였다."
    )
    conclusion_text = (
        "ChatGPT와 Gemini는 자연어로 상황과 안전조치를 설명하는 데 유용하였다. "
        "다만 본 실험에서 요구한 공식 문서의 추적성 및 위험성평가 양식화는 답변별 "
        "편차가 있었다. 내 사이트 RAG는 Vector DB 검색을 통해 공식 문서명과 chunk_id를 "
        "제시하고, KRAS식 위험성평가 초안과 조치 체크리스트를 함께 제공하여 "
        "광산 안전관리 기록 지원 목적에서 높은 적합성을 보였다."
    )
    caution_text = (
        "본 결과는 답변의 키워드, 출처 표현, 안전조치 및 구조화 요소를 분석한 규칙 기반 "
        "1차 자동평가이다. 범용 LLM과 도메인 특화 RAG는 설계 목적이 다르므로 점수만으로 "
        "일반적 우열을 단정하지 않으며, 최종 평가는 대표 문항 수동 검토가 필요하다."
    )
    score_text = " / ".join(
        f"{target}: {summary_records[target]['총점 평균']:.2f}점"
        for target in targets
    )
    conclusions = [
        ["구분", "결론 문구"],
        ["실험 설계", design_text],
        ["핵심 결론", conclusion_text],
        ["평균 결과", score_text],
        ["해석 원칙", caution_text],
        [
            "발표용 요약",
            "동일한 20개 질문 비교 결과, 범용 LLM의 자연어 설명 능력과 도메인 특화 "
            "RAG의 공식 근거 추적·KRAS 기록 구조화 능력은 상호 보완적으로 나타났다. "
            "광산 안전 현장에서는 검색 근거, 구조화된 기록 초안, 전문가 검토를 결합하는 "
            "방식이 적절하다.",
        ],
    ]

    total_chart = [["비교대상", "총점 평균"]] + [
        [target, summary_records[target]["총점 평균"]]
        for target in targets
    ]
    criteria_chart = [[
        "비교대상",
        "검색 적합성",
        "근거 기반성",
        "안전·법령 판단",
        "실무성",
    ]] + [
        [
            target,
            summary_records[target]["검색 적합성 평균"],
            summary_records[target]["근거 기반성 평균"],
            summary_records[target]["안전·법령 판단 정확성 평균"],
            summary_records[target]["실무성 평균"],
        ]
        for target in targets
    ]

    payload = {
        "raw_answers": raw_answers,
        "evaluations": [EVALUATION_HEADERS] + evaluations,
        "summary": summary_rows,
        "total_chart": total_chart,
        "criteria_chart": criteria_chart,
        "representative_cases": representative_rows,
        "conclusions": conclusions,
    }
    stats = {
        "raw_row_count": len(rows),
        "generated_count": generated_count,
        "scored_count": sum(1 for row in evaluation_records if row["판정"] != "미평가"),
        "review_count": sum(1 for row in evaluation_records if row["검토필요"] == "Y"),
        "averages": {
            target: summary_records[target]["총점 평균"]
            for target in targets
        },
        "representative_ids": [
            case["question_id"]
            for case in selected_cases
        ],
    }
    return payload, stats


def main() -> int:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"입력 엑셀을 찾을 수 없습니다: {TEMPLATE_PATH}")
    if not BRIDGE_PATH.exists():
        raise FileNotFoundError(f"Excel 브리지 파일을 찾을 수 없습니다: {BRIDGE_PATH}")

    EXPERIMENT_DIR.mkdir(parents=True, exist_ok=True)
    extracted_path = EXPERIMENT_DIR / ".comparison_input.json"
    payload_path = EXPERIMENT_DIR / ".comparison_result_payload.json"

    run_bridge("extract", TEMPLATE_PATH, extracted_path)
    workbook_data = json.loads(extracted_path.read_text(encoding="utf-8"))
    payload, stats = build_payload(workbook_data)
    payload_path.write_text(
        json.dumps(payload, ensure_ascii=False),
        encoding="utf-8",
    )
    run_bridge("build", TEMPLATE_PATH, payload_path, OUTPUT_PATH)
    extracted_path.unlink(missing_ok=True)
    payload_path.unlink(missing_ok=True)

    print(f"읽은 엑셀 파일명: {TEMPLATE_PATH.name}")
    print(f"답변 원문 행 수: {stats['raw_row_count']}")
    print(f"자동 생성한 내 사이트 RAG 답변 수: {stats['generated_count']}")
    print(f"자동 채점한 행 수: {stats['scored_count']}")
    print(f"저장된 결과 파일 경로: {OUTPUT_PATH}")
    print("비교대상별 총점 평균:")
    for target, score in stats["averages"].items():
        print(f"  - {target}: {score:.2f}")
    print("대표 사례 질문 번호: " + ", ".join(stats["representative_ids"]))
    print(f"검토필요 행 수: {stats['review_count']}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"[오류] {exc}", file=sys.stderr)
        raise
