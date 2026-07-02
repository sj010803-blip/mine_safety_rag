from pathlib import Path
import json
import sys

import chromadb
from sentence_transformers import SentenceTransformer


# ==============================
# 경로 설정
# ==============================
ROOT_DIR = Path(__file__).resolve().parents[1]

CHUNKS_PATH = ROOT_DIR / "08_chunks" / "chunks.jsonl"
VECTOR_DB_DIR = ROOT_DIR / "10_vector_db"

COLLECTION_NAME = "mine_safety_docs"

EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"


# ==============================
# chunk 본문 찾기
# ==============================
TEXT_KEYS = [
    "text",
    "chunk_text",
    "content",
    "page_content",
    "document",
]


def find_text(item: dict) -> str:
    """chunks.jsonl 안에서 실제 본문 텍스트를 찾는 함수"""
    for key in TEXT_KEYS:
        value = item.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()

    metadata = item.get("metadata")
    if isinstance(metadata, dict):
        for key in TEXT_KEYS:
            value = metadata.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()

    return ""


def clean_metadata_value(value):
    """ChromaDB metadata에 넣을 수 있는 값으로 변환"""
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return json.dumps(value, ensure_ascii=False)


def make_metadata(item: dict, index: int) -> dict:
    """chunk별 metadata 생성"""
    metadata = {}

    raw_metadata = item.get("metadata")
    if isinstance(raw_metadata, dict):
        for key, value in raw_metadata.items():
            metadata[key] = clean_metadata_value(value)

    for key, value in item.items():
        if key in TEXT_KEYS:
            continue
        if key in ["metadata", "embedding", "embeddings"]:
            continue
        metadata[key] = clean_metadata_value(value)

    metadata["chunk_index"] = index

    if "source" not in metadata:
        for key in ["file_name", "doc_name", "title", "source_file"]:
            if key in metadata and metadata[key]:
                metadata["source"] = metadata[key]
                break

    if "source" not in metadata:
        metadata["source"] = "출처 정보 없음"

    return metadata


def make_id(item: dict, index: int, used_ids: set) -> str:
    """ChromaDB에 넣을 고유 ID 생성"""
    candidate = (
        item.get("chunk_id")
        or item.get("id")
        or item.get("doc_id")
        or f"chunk_{index:05d}"
    )

    candidate = str(candidate)

    if candidate not in used_ids:
        used_ids.add(candidate)
        return candidate

    new_candidate = f"{candidate}_{index:05d}"
    used_ids.add(new_candidate)
    return new_candidate


def load_chunks():
    """chunks.jsonl 파일 불러오기"""
    if not CHUNKS_PATH.exists():
        print(f"[오류] chunks.jsonl 파일을 찾을 수 없습니다.")
        print(f"확인 경로: {CHUNKS_PATH}")
        sys.exit(1)

    chunks = []

    with open(CHUNKS_PATH, "r", encoding="utf-8") as f:
        for index, line in enumerate(f):
            line = line.strip()

            if not line:
                continue

            try:
                item = json.loads(line)
            except json.JSONDecodeError:
                print(f"[경고] {index + 1}번째 줄 JSON 읽기 실패. 건너뜁니다.")
                continue

            text = find_text(item)

            if not text:
                print(f"[경고] {index + 1}번째 chunk에 텍스트가 없습니다. 건너뜁니다.")
                continue

            chunks.append((index, item, text))

    return chunks


def main():
    print("=" * 70)
    print("광산 안전 AI 에이전트 Vector DB 구축 시작")
    print("=" * 70)

    print(f"[정보] 프로젝트 폴더: {ROOT_DIR}")
    print(f"[정보] chunk 파일: {CHUNKS_PATH}")
    print(f"[정보] Vector DB 저장 폴더: {VECTOR_DB_DIR}")

    VECTOR_DB_DIR.mkdir(parents=True, exist_ok=True)

    # 1. chunk 불러오기
    chunks = load_chunks()

    if not chunks:
        print("[오류] 불러온 chunk가 없습니다.")
        sys.exit(1)

    print(f"[정보] 불러온 chunk 수: {len(chunks)}개")

    # 2. 임베딩 모델 로드
    print("[정보] 임베딩 모델을 불러오는 중입니다.")
    print("[정보] 첫 실행 시 모델 다운로드 때문에 시간이 걸릴 수 있습니다.")

    model = SentenceTransformer(EMBEDDING_MODEL_NAME)

    # 3. ChromaDB 클라이언트 생성
    client = chromadb.PersistentClient(path=str(VECTOR_DB_DIR))

    # 기존 collection 삭제 후 새로 생성
    try:
        client.delete_collection(name=COLLECTION_NAME)
        print(f"[정보] 기존 collection 삭제 완료: {COLLECTION_NAME}")
    except Exception:
        print(f"[정보] 기존 collection 없음. 새로 생성합니다: {COLLECTION_NAME}")

    collection = client.get_or_create_collection(
        name=COLLECTION_NAME,
        metadata={
            "description": "광산 안전 지침 및 중대재해처벌법 대응 RAG Vector DB",
            "project": "mine_safety_rag",
        },
    )

    ids = []
    documents = []
    metadatas = []
    used_ids = set()

    for index, item, text in chunks:
        chunk_id = make_id(item, index, used_ids)
        metadata = make_metadata(item, index)

        ids.append(chunk_id)
        documents.append(text)
        metadatas.append(metadata)

    # 4. 문서 임베딩 생성 및 저장
    batch_size = 64
    total = len(documents)

    print("[정보] chunk 임베딩 생성 및 ChromaDB 저장 시작")

    for start in range(0, total, batch_size):
        end = min(start + batch_size, total)

        batch_docs = documents[start:end]
        batch_ids = ids[start:end]
        batch_metas = metadatas[start:end]

        embeddings = model.encode(
            batch_docs,
            normalize_embeddings=True,
            show_progress_bar=False,
        ).tolist()

        collection.add(
            ids=batch_ids,
            documents=batch_docs,
            metadatas=batch_metas,
            embeddings=embeddings,
        )

        print(f"  - 저장 완료: {end}/{total}")

    print("=" * 70)
    print("[완료] Vector DB 구축 완료")
    print(f"[완료] 저장된 chunk 수: {collection.count()}개")
    print(f"[완료] 저장 위치: {VECTOR_DB_DIR}")
    print("=" * 70)

    # 5. 테스트 검색
    print("\n[테스트 검색]")
    query = "갱내 메탄가스 기준은 어떻게 확인해야 해?"

    query_embedding = model.encode(
        [query],
        normalize_embeddings=True,
        show_progress_bar=False,
    ).tolist()

    results = collection.query(
        query_embeddings=query_embedding,
        n_results=5,
        include=["documents", "metadatas", "distances"],
    )

    print(f"질문: {query}")

    docs = results.get("documents", [[]])[0]
    metas = results.get("metadatas", [[]])[0]
    distances = results.get("distances", [[]])[0]

    for i, doc in enumerate(docs):
        meta = metas[i] if i < len(metas) else {}
        distance = distances[i] if i < len(distances) else None

        source = (
            meta.get("source")
            or meta.get("file_name")
            or meta.get("doc_name")
            or meta.get("title")
            or "출처 정보 없음"
        )

        print("\n" + "-" * 70)
        print(f"검색 결과 {i + 1}")
        print(f"출처: {source}")
        print(f"거리값: {distance}")
        print("본문 미리보기:")
        print(doc[:500].replace("\n", " "))

    print("\n[완료] 테스트 검색까지 완료되었습니다.")


if __name__ == "__main__":
    main()