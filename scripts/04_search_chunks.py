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
    terms = [query] + words
    # 중복 제거
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
            # 긴 검색어가 직접 들어가면 더 높은 점수
            score += count * (len(t) + 3)

    return score

def short_preview(text, max_len=700):
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > max_len:
        return text[:max_len] + "..."
    return text

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

def main():
    print("=" * 60)
    print("Mine Safety RAG simple chunk search")
    print("=" * 60)

    if not CHUNK_FILE.exists():
        print("ERROR: chunks.jsonl 파일이 없습니다.")
        print("먼저 03_make_chunks.py를 실행해야 합니다.")
        print("찾는 위치:", CHUNK_FILE)
        return

    chunks = load_chunks()
    print("Loaded chunks:", len(chunks))
    print()

    while True:
        query = input("검색어 또는 질문 입력, 종료하려면 q 입력: ").strip()

        if query.lower() in ["q", "quit", "exit"]:
            print("검색 종료")
            break

        if not query:
            continue

        results = search(chunks, query, top_k=5)

        print()
        print("-" * 60)
        print("Query:", query)
        print("Results:", len(results))
        print("-" * 60)

        if not results:
            print("검색 결과가 없습니다. 다른 단어로 검색해보세요.")
            print()
            continue

        for rank, (score, chunk) in enumerate(results, start=1):
            print(f"\n[{rank}] score={score}")
            print("source_file:", chunk.get("source_file"))
            print("chunk_id:", chunk.get("chunk_id"))
            print("chunk_index:", chunk.get("chunk_index"))
            print("preview:")
            print(short_preview(chunk.get("text", "")))
            print("-" * 60)

        print()

if __name__ == "__main__":
    main()
