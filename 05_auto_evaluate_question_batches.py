from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from pathlib import Path
from typing import Any

import chromadb
from sentence_transformers import SentenceTransformer

ROOT_DIR = Path(__file__).resolve().parent
VECTOR_DB_DIR = ROOT_DIR / "10_vector_db"
SCENARIO_DIR = ROOT_DIR / "02_질문시나리오"
CHUNKS_PATH = ROOT_DIR / "08_chunks" / "chunks.jsonl"
COLLECTION_NAME = "mine_safety_docs"
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"

FULL_SCENARIO_PATH = SCENARIO_DIR / "question_scenarios_100.tsv"
FIRST35_OUTPUT_PATH = ROOT_DIR / "09_answer_tests" / "auto_eval_Q031_Q065.tsv"
SECOND35_OUTPUT_PATH = ROOT_DIR / "09_answer_tests" / "auto_eval_Q066_Q100.tsv"
SUMMARY_OUTPUT_PATH = ROOT_DIR / "09_answer_tests" / "auto_eval_100_summary.tsv"

SAFETY_KEYWORDS = [
    "작업중지",
    "대피",
    "접근금지",
    "보고",
    "보호구",
    "환기",
    "가스측정",
    "차단",
    "잠금표지",
    "위험성평가",
    "사고보고",
    "응급조치",
    "중대재해",
    "재발방지",
    "기록",
    "작업허가",
    "출입통제",
    "재개",
    "점검",
]

PRACTICAL_KEYWORDS = [
    "즉시",
    "확인",
    "점검",
    "기록",
    "보고",
    "재개",
    "공유",
    "통제",
    "대피",
    "작업중지",
    "출입",
    "보호구",
    "재측정",
    "조치",
]

RISKY_PHRASES = [
    "확정적 법률 해석",
    "법령 해석은 확정",
    "무조건 허용",
    "무조건 가능",
    "반드시 정답",
]


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def normalize_question_no(value: Any) -> str:
    text = str(value or "").strip().upper()
    if text.startswith("Q"):
        text = text[1:]
    try:
        return str(int(text))
    except ValueError:
        return text


def load_scenarios(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        raise FileNotFoundError(f"질문 시나리오 파일이 없습니다: {path}")
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter="\t"))


def read_chunks_count() -> int:
    if not CHUNKS_PATH.exists():
        return 0
    with open(CHUNKS_PATH, "r", encoding="utf-8") as f:
        return sum(1 for _ in f if _.strip())


def load_collection_and_model():
    client = chromadb.PersistentClient(path=str(VECTOR_DB_DIR))
    collection = client.get_collection(name=COLLECTION_NAME)
    model = SentenceTransformer(EMBEDDING_MODEL_NAME)
    return collection, model


def search_vector_db(collection, model, question: str, top_k: int = 5) -> list[dict[str, Any]]:
    query_embedding = model.encode([question], normalize_embeddings=True, show_progress_bar=False).tolist()
    results = collection.query(query_embeddings=query_embedding, n_results=max(top_k, 5), include=["documents", "metadatas", "distances"])
    docs = results.get("documents", [[]])[0]
    metas = results.get("metadatas", [[]])[0]
    distances = results.get("distances", [[]])[0]

    cleaned_results = []
    for idx, doc in enumerate(docs):
        meta = metas[idx] if idx < len(metas) and isinstance(metas[idx], dict) else {}
        distance = distances[idx] if idx < len(distances) else None
        text = clean_text(doc or "")
        if not text:
            continue
        source = str(meta.get("source") or meta.get("file_name") or meta.get("doc_name") or "출처 없음")
        chunk_id = str(meta.get("chunk_id") or meta.get("id") or meta.get("doc_id") or meta.get("index") or "정보 없음")
        cleaned_results.append(
            {
                "rank": idx + 1,
                "text": text,
                "source": source,
                "chunk_id": chunk_id,
                "distance": distance,
            }
        )
    return cleaned_results[:top_k]


def build_answer_draft(question: str, results: list[dict[str, Any]]) -> str:
    if not results:
        return "검색 결과가 부족하여 자동 평가용 초안 생성이 제한됩니다."

    top_texts = []
    for item in results[:3]:
        snippet = clean_text(item["text"])
        if len(snippet) > 220:
            snippet = snippet[:220].rstrip() + "..."
        top_texts.append(f"- {item['source']} | chunk_id {item['chunk_id']} | {snippet}")

    return "\n".join([
        "1차 자동 평가용 답변 초안",
        f"질문: {question}",
        "근거 요약:",
        *top_texts,
        "현장 조치 기준: 작업 전 위험요인 확인, 작업중지·대피 기준, 보고·기록 절차, 보호구·환기·가스측정 여부 확인",
    ])


def calculate_scores(question: str, expected_keywords: list[str], results: list[dict[str, Any]], answer_draft: str) -> tuple[dict[str, int], str, bool]:
    if not results:
        return {
            "검색_적합성": 0,
            "근거_기반성": 0,
            "안전법령_판단정확성": 0,
            "실무성": 0,
        }, "검색 결과가 없어 자동 평가가 제한됩니다.", True

    lower_question = question.lower()
    combined_text = " ".join([item["text"].lower() for item in results[:3]])
    expected_text = " ".join(expected_keywords).lower()

    keyword_hits = sum(1 for kw in expected_keywords if kw.lower() in combined_text)
    top3_related = sum(1 for item in results[:3] if any(kw.lower() in item["text"].lower() for kw in expected_keywords[:3]))

    search_score = 6
    if keyword_hits >= 3:
        search_score += 8
    elif keyword_hits >= 1:
        search_score += 4
    if top3_related >= 2:
        search_score += 6
    elif top3_related >= 1:
        search_score += 3
    if any(token in lower_question for token in ["작업전", "작업 전", "작업중지", "재개", "보호구", "환기", "가스", "전기", "차량", "위험성평가", "응급", "사고"]):
        search_score += 2
    search_score = max(0, min(25, search_score))

    evidence_score = 6
    if any(item["source"] for item in results[:3]):
        evidence_score += 4
    if any(kw.lower() in answer_draft.lower() for kw in ["chunk_id", "근거", "작업중지", "대피", "보호구", "기록", "보고"]):
        evidence_score += 6
    if keyword_hits >= 2:
        evidence_score += 4
    evidence_score = max(0, min(25, evidence_score))

    safety_score = 6
    if any(keyword.lower() in answer_draft.lower() for keyword in SAFETY_KEYWORDS):
        safety_score += 8
    if any(keyword.lower() in answer_draft.lower() for keyword in ["작업중지", "대피", "보고", "환기", "차단", "잠금표지", "재개", "가스측정"]):
        safety_score += 5
    if any(phrase.lower() in answer_draft.lower() for phrase in RISKY_PHRASES):
        safety_score -= 8
    safety_score = max(0, min(25, safety_score))

    practicality_score = 6
    if any(keyword.lower() in answer_draft.lower() for keyword in PRACTICAL_KEYWORDS):
        practicality_score += 8
    if any(keyword.lower() in answer_draft.lower() for keyword in ["즉시", "확인", "점검", "기록", "보고", "재개", "공유", "통제"]):
        practicality_score += 5
    practicality_score = max(0, min(25, practicality_score))

    scores = {
        "검색_적합성": int(search_score),
        "근거_기반성": int(evidence_score),
        "안전법령_판단정확성": int(safety_score),
        "실무성": int(practicality_score),
    }

    total = sum(scores.values())
    review_needed = total < 80 or len(results) < 3 or keyword_hits < 2 or search_score < 12
    reason_parts = []
    if total < 80:
        reason_parts.append("총점이 80점 미만")
    if len(results) < 3:
        reason_parts.append("검색 결과 부족")
    if keyword_hits < 2:
        reason_parts.append("예상 키워드 매칭 약함")
    if search_score < 12:
        reason_parts.append("상위 근거 연관성 약함")
    if not reason_parts:
        reason = "상위 근거와 실무 조치 요소가 비교적 잘 반영됨"
    else:
        reason = "; ".join(reason_parts)
    return scores, reason, review_needed


def calculate_judgment(total_score: int) -> str:
    if total_score >= 90:
        return "매우 우수"
    if total_score >= 80:
        return "우수"
    if total_score >= 70:
        return "보통"
    if total_score >= 60:
        return "보완 필요"
    return "미흡"


def select_scenarios(scenarios: list[dict[str, str]], start: int | None = None, end: int | None = None) -> list[dict[str, str]]:
    if start is None and end is None:
        return scenarios
    if start is None:
        start = 0
    if end is None:
        end = len(scenarios)
    return scenarios[start:end]


def write_tsv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = [
        "question_id",
        "category",
        "question",
        "expected_keywords",
        "retrieved_chunk_ids",
        "retrieved_sources",
        "retrieved_distances",
        "answer_draft",
        "검색_적합성",
        "근거_기반성",
        "안전법령_판단정확성",
        "실무성",
        "총점",
        "판정",
        "검토필요",
        "평가_근거",
        "평가방식",
    ]
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers, delimiter="\t", lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def evaluate_batch(scenarios: list[dict[str, str]], output_path: Path, batch_name: str) -> list[dict[str, Any]]:
    collection, model = load_collection_and_model()
    rows: list[dict[str, Any]] = []
    for row in scenarios:
        question = clean_text(row.get("질문 시나리오", ""))
        category = clean_text(row.get("분류", ""))
        no_text = normalize_question_no(row.get("번호", ""))
        expected_keywords = [kw.strip() for kw in str(row.get("기대 검색 문서", "")).split(",") if kw.strip()]
        if not expected_keywords:
            expected_keywords = [category]
        results = search_vector_db(collection, model, question, top_k=5)
        answer_draft = build_answer_draft(question, results)
        scores, reason, review_needed = calculate_scores(question, expected_keywords, results, answer_draft)
        total_score = sum(scores.values())
        judgment = calculate_judgment(total_score)
        rows.append(
            {
                "question_id": f"Q{int(no_text):02d}" if no_text.isdigit() else no_text,
                "category": category,
                "question": question,
                "expected_keywords": ",".join(expected_keywords),
                "retrieved_chunk_ids": ",".join(str(item["chunk_id"]) for item in results),
                "retrieved_sources": ",".join(str(item["source"]) for item in results),
                "retrieved_distances": ",".join(str(item["distance"]) for item in results),
                "answer_draft": answer_draft,
                "검색_적합성": scores["검색_적합성"],
                "근거_기반성": scores["근거_기반성"],
                "안전법령_판단정확성": scores["안전법령_판단정확성"],
                "실무성": scores["실무성"],
                "총점": total_score,
                "판정": judgment,
                "검토필요": "Y" if review_needed else "N",
                "평가_근거": reason,
                "평가방식": "1차 자동 평가(rule-based)",
            }
        )
    write_tsv(output_path, rows)
    return rows


def build_summary(rows: list[dict[str, Any]], output_path: Path) -> None:
    judgment_counts = {label: 0 for label in ["매우 우수", "우수", "보통", "보완 필요", "미흡"]}
    review_count = 0
    total_score = 0
    for row in rows:
        total_score += int(row.get("총점", 0))
        judgment_counts[str(row.get("판정", "미흡"))] = judgment_counts.get(str(row.get("판정", "미흡")), 0) + 1
        if str(row.get("검토필요", "N")).upper() == "Y":
            review_count += 1
    count = len(rows)
    avg_score = round(total_score / count, 2) if count else 0.0

    summary_rows = [{
        "question_id": "Q001-Q100",
        "category": "요약",
        "question": "전체 자동 평가 요약",
        "expected_keywords": "",
        "retrieved_chunk_ids": "",
        "retrieved_sources": "",
        "retrieved_distances": "",
        "answer_draft": "",
        "검색_적합성": "",
        "근거_기반성": "",
        "안전법령_판단정확성": "",
        "실무성": "",
        "총점": avg_score,
        "판정": "",
        "검토필요": review_count,
        "평가_근거": f"평가 질문 수={count}; 평균점수={avg_score}; 매우 우수={judgment_counts['매우 우수']}; 우수={judgment_counts['우수']}; 보통={judgment_counts['보통']}; 보완 필요={judgment_counts['보완 필요']}; 미흡={judgment_counts['미흡']}; 검토필요 Y={review_count}",
        "평가방식": "1차 자동 평가(rule-based)",
    }]
    write_tsv(output_path, summary_rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="추가 질문 자동 평가")
    parser.add_argument("--batch", choices=["first35", "second35", "all100"], required=True)
    args = parser.parse_args()

    chunk_count = read_chunks_count()
    print(f"chunks.jsonl 기준 문서 수: {chunk_count}")

    all_scenarios = load_scenarios(FULL_SCENARIO_PATH)

    if args.batch == "first35":
        scenarios = select_scenarios(all_scenarios, start=30, end=65)
        result_rows = evaluate_batch(scenarios, FIRST35_OUTPUT_PATH, "first35")
        avg_score = round(sum(int(row["총점"]) for row in result_rows) / len(result_rows), 2) if result_rows else 0.0
        print(f"평가한 질문 수: {len(result_rows)}")
        print(f"평균 점수: {avg_score}")
        print(f"저장 파일: {FIRST35_OUTPUT_PATH}")
    elif args.batch == "second35":
        scenarios = select_scenarios(all_scenarios, start=65, end=100)
        result_rows = evaluate_batch(scenarios, SECOND35_OUTPUT_PATH, "second35")
        avg_score = round(sum(int(row["총점"]) for row in result_rows) / len(result_rows), 2) if result_rows else 0.0
        print(f"평가한 질문 수: {len(result_rows)}")
        print(f"평균 점수: {avg_score}")
        print(f"저장 파일: {SECOND35_OUTPUT_PATH}")
    else:
        first_rows = evaluate_batch(select_scenarios(all_scenarios, start=30, end=65), FIRST35_OUTPUT_PATH, "first35")
        second_rows = evaluate_batch(select_scenarios(all_scenarios, start=65, end=100), SECOND35_OUTPUT_PATH, "second35")
        all_rows = first_rows + second_rows
        build_summary(all_rows, SUMMARY_OUTPUT_PATH)
        avg_score = round(sum(int(row["총점"]) for row in all_rows) / len(all_rows), 2) if all_rows else 0.0
        review_count = sum(1 for row in all_rows if str(row.get("검토필요", "N")).upper() == "Y")
        judgment_counts = {label: 0 for label in ["매우 우수", "우수", "보통", "보완 필요", "미흡"]}
        for row in all_rows:
            judgment_counts[str(row.get("판정", "미흡"))] = judgment_counts.get(str(row.get("판정", "미흡")), 0) + 1
        print(f"평가한 질문 수: {len(all_rows)}")
        print(f"평균 점수: {avg_score}")
        print(f"매우 우수 / 우수 / 보통 / 보완 필요 / 미흡: {judgment_counts['매우 우수']} / {judgment_counts['우수']} / {judgment_counts['보통']} / {judgment_counts['보완 필요']} / {judgment_counts['미흡']}")
        print(f"검토필요 Y 개수: {review_count}")
        print(f"저장 파일: {SUMMARY_OUTPUT_PATH}")


if __name__ == "__main__":
    main()
