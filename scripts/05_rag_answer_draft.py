from pathlib import Path
import json
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE_DIR = Path(__file__).resolve().parents[1]
CHUNK_FILE = BASE_DIR / "08_chunks" / "chunks.jsonl"
OUTPUT_DIR = BASE_DIR / "09_answer_tests"
OUTPUT_DIR.mkdir(exist_ok=True)

def load_chunks():
    chunks = []
    with open(CHUNK_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                chunks.append(json.loads(line))
    return chunks

def make_terms(query):
    query = query.strip()
    words = re.findall(r"[가-힣A-Za-z0-9_.%-]+", query)

    extra_terms = []

    # 질문 유형별 관련 키워드 보강
    if "메탄" in query:
        extra_terms += ["메탄가스", "갑종탄광", "석탄광산", "가스측정기"]
    if "산소" in query:
        extra_terms += ["산소함유량", "갱내의 공기", "19%"]
    if "중대" in query:
        extra_terms += ["중대산업재해", "중대재해", "경영책임자", "안전보건관리체계"]
    if "위험성" in query:
        extra_terms += ["위험성평가", "유해", "위험요인", "근로자 참여", "TBM"]
    if "발파" in query:
        extra_terms += ["발파", "화약", "폭발", "낙반", "비산"]
    if "굴착" in query or "갱내" in query:
        extra_terms += ["굴착", "갱내", "통기", "유해가스", "작업장"]

    terms = [query] + words + extra_terms

    result = []
    for term in terms:
        term = term.strip()
        if term and term not in result:
            result.append(term)

    return result

def score_chunk(chunk_text, terms):
    text = chunk_text.lower()
    score = 0

    for term in terms:
        t = term.lower()
        if not t:
            continue

        count = text.count(t)
        if count:
            score += count * (len(t) + 3)

    return score

def search(chunks, query, top_k=5):
    terms = make_terms(query)

    results = []
    for chunk in chunks:
        text = chunk.get("text", "")
        score = score_chunk(text, terms)

        if score > 0:
            results.append((score, chunk))

    results.sort(key=lambda x: x[0], reverse=True)
    return results[:top_k]

def clean_text(text, max_len=900):
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > max_len:
        text = text[:max_len] + "..."
    return text

def make_answer_draft(query, results):
    lines = []
    lines.append("=" * 70)
    lines.append("질문")
    lines.append("=" * 70)
    lines.append(query)
    lines.append("")

    if not results:
        lines.append("검색된 근거 문서가 없습니다.")
        lines.append("검색어를 더 짧게 바꿔보세요. 예: 메탄가스, 위험성평가, 중대산업재해")
        return "\n".join(lines)

    lines.append("=" * 70)
    lines.append("답변 초안")
    lines.append("=" * 70)
    lines.append("아래 답변은 AI가 최종 판단한 법률 자문이 아니라, 수집한 문서 조각을 바탕으로 만든 초안입니다.")
    lines.append("최종 발표나 보고서에는 반드시 원문 문서의 조항과 페이지를 다시 확인해야 합니다.")
    lines.append("")

    lines.append("핵심 요약:")
    lines.append("- 이 질문과 관련된 근거 문서를 검색했습니다.")
    lines.append("- 아래의 근거 문서 내용을 기준으로 현장 점검 항목, 법적 근거, 위험요인을 정리할 수 있습니다.")
    lines.append("- 실제 RAG 에이전트에서는 이 검색 결과를 LLM에 넣어 자연스러운 답변으로 바꾸게 됩니다.")
    lines.append("")

    lines.append("=" * 70)
    lines.append("검색된 근거 문서 TOP 5")
    lines.append("=" * 70)

    for rank, (score, chunk) in enumerate(results, start=1):
        lines.append("")
        lines.append(f"[근거 {rank}] score={score}")
        lines.append(f"source_file: {chunk.get('source_file')}")
        lines.append(f"chunk_id: {chunk.get('chunk_id')}")
        lines.append(f"chunk_index: {chunk.get('chunk_index')}")
        lines.append("내용 미리보기:")
        lines.append(clean_text(chunk.get("text", ""), max_len=900))

    return "\n".join(lines)

def save_answer(query, answer_text):
    safe_name = re.sub(r"[^가-힣A-Za-z0-9_-]+", "_", query).strip("_")
    if not safe_name:
        safe_name = "answer"
    safe_name = safe_name[:40]

    output_path = OUTPUT_DIR / f"answer_{safe_name}.txt"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(answer_text)

    return output_path

def main():
    print("=" * 70)
    print("Mine Safety RAG answer draft test")
    print("=" * 70)

    if not CHUNK_FILE.exists():
        print("ERROR: chunks.jsonl 파일이 없습니다.")
        print("먼저 03_make_chunks.py를 실행해야 합니다.")
        print("찾는 위치:", CHUNK_FILE)
        return

    chunks = load_chunks()
    print("Loaded chunks:", len(chunks))
    print()
    print("예시 질문:")
    print("- 갱내 메탄가스 기준은 어떻게 확인해야 해?")
    print("- 중대산업재해가 발생하면 어떤 의무가 중요해?")
    print("- 위험성평가는 어떤 절차로 해야 해?")
    print("- 갱내 산소함유량 기준은 얼마야?")
    print()
    print("종료하려면 q 입력")
    print()

    while True:
        query = input("질문 입력: ").strip()

        if query.lower() in ["q", "quit", "exit"]:
            print("종료")
            break

        if not query:
            continue

        results = search(chunks, query, top_k=5)
        answer_text = make_answer_draft(query, results)
        output_path = save_answer(query, answer_text)

        print()
        print(answer_text)
        print()
        print("저장 위치:", output_path)
        print()

if __name__ == "__main__":
    main()
