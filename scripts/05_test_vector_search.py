from pathlib import Path
import csv

import chromadb
from sentence_transformers import SentenceTransformer


# ==============================
# 경로 설정
# ==============================
ROOT_DIR = Path(__file__).resolve().parents[1]

VECTOR_DB_DIR = ROOT_DIR / "10_vector_db"
RESULT_DIR = ROOT_DIR / "09_answer_tests"
RESULT_CSV = RESULT_DIR / "vector_search_test_results.csv"

COLLECTION_NAME = "mine_safety_docs"
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"


# ==============================
# 테스트 질문 목록
# ==============================
TEST_QUESTIONS = [
    "갱내 메탄가스 기준은 어떻게 확인해야 해?",
    "발파 작업 전에 현장 관리자가 확인해야 할 안전 체크리스트 알려줘",
    "갱내 작업 전에 환기 상태는 어떻게 점검해야 해?",
    "낙반 위험이 있을 때 작업을 중지해야 하는 기준은 뭐야?",
    "중대재해가 발생하면 현장 관리자는 먼저 무엇을 해야 해?",
    "광산에서 유해가스가 발생했을 때 필요한 조치사항 알려줘",
    "위험성평가는 작업 전에 어떤 식으로 해야 해?",
    "광산 전기설비 점검 시 감전 예방 조치는 뭐가 있어?",
    "작업자가 보호구를 착용하지 않았을 때 관리자가 해야 할 조치는?",
    "중대재해처벌법상 경영책임자의 안전보건 확보의무를 요약해줘",
]


def get_source(meta: dict) -> str:
    """검색 결과의 출처명 찾기"""
    return (
        meta.get("source")
        or meta.get("file_name")
        or meta.get("doc_name")
        or meta.get("title")
        or meta.get("source_file")
        or "출처 정보 없음"
    )


def clean_text(text: str, limit: int = 350) -> str:
    """본문 미리보기 정리"""
    text = text.replace("\n", " ").replace("\t", " ")
    text = " ".join(text.split())

    if len(text) > limit:
        return text[:limit] + "..."

    return text


def main():
    print("=" * 70)
    print("광산 안전 AI 에이전트 Vector DB 검색 테스트 시작")
    print("=" * 70)

    print(f"[정보] 프로젝트 폴더: {ROOT_DIR}")
    print(f"[정보] Vector DB 폴더: {VECTOR_DB_DIR}")
    print(f"[정보] 결과 저장 위치: {RESULT_CSV}")

    if not VECTOR_DB_DIR.exists():
        print("[오류] 10_vector_db 폴더가 없습니다.")
        print("먼저 scripts\\04_build_vector_db.py를 실행해서 Vector DB를 만들어야 합니다.")
        return

    RESULT_DIR.mkdir(parents=True, exist_ok=True)

    # 1. 임베딩 모델 불러오기
    print("[정보] 임베딩 모델을 불러오는 중입니다.")
    model = SentenceTransformer(EMBEDDING_MODEL_NAME)

    # 2. ChromaDB 불러오기
    client = chromadb.PersistentClient(path=str(VECTOR_DB_DIR))

    try:
        collection = client.get_collection(name=COLLECTION_NAME)
    except Exception as e:
        print("[오류] ChromaDB collection을 찾을 수 없습니다.")
        print(f"collection 이름: {COLLECTION_NAME}")
        print("먼저 scripts\\04_build_vector_db.py를 다시 실행해보세요.")
        print(e)
        return

    print(f"[정보] Vector DB에 저장된 chunk 수: {collection.count()}개")

    rows = []

    # 3. 질문별 검색 테스트
    for q_idx, question in enumerate(TEST_QUESTIONS, start=1):
        print("\n" + "=" * 70)
        print(f"[질문 {q_idx}] {question}")
        print("=" * 70)

        query_embedding = model.encode(
            [question],
            normalize_embeddings=True,
            show_progress_bar=False,
        ).tolist()

        results = collection.query(
            query_embeddings=query_embedding,
            n_results=5,
            include=["documents", "metadatas", "distances"],
        )

        docs = results.get("documents", [[]])[0]
        metas = results.get("metadatas", [[]])[0]
        distances = results.get("distances", [[]])[0]

        if not docs:
            print("[경고] 검색 결과가 없습니다.")
            rows.append({
                "question_no": q_idx,
                "question": question,
                "rank": "",
                "source": "검색 결과 없음",
                "distance": "",
                "preview": "",
            })
            continue

        for rank, doc in enumerate(docs, start=1):
            meta = metas[rank - 1] if rank - 1 < len(metas) else {}
            distance = distances[rank - 1] if rank - 1 < len(distances) else ""

            source = get_source(meta)
            preview = clean_text(doc)

            print(f"\n검색 결과 {rank}")
            print(f"출처: {source}")
            print(f"거리값: {distance}")
            print(f"본문 미리보기: {preview}")

            rows.append({
                "question_no": q_idx,
                "question": question,
                "rank": rank,
                "source": source,
                "distance": distance,
                "preview": preview,
            })

    # 4. CSV 저장
    with open(RESULT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        fieldnames = ["question_no", "question", "rank", "source", "distance", "preview"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)

        writer.writeheader()
        writer.writerows(rows)

    print("\n" + "=" * 70)
    print("[완료] Vector DB 검색 테스트 완료")
    print(f"[완료] 결과 CSV 저장: {RESULT_CSV}")
    print("=" * 70)


if __name__ == "__main__":
    main()