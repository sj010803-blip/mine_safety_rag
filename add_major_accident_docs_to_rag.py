from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parent
RAW_ADDED_DIR = PROJECT_ROOT / "01_raw_docs" / "중대재해처벌법_추가자료"
ORIGINAL_CHUNKS_PATH = PROJECT_ROOT / "08_chunks" / "chunks.jsonl"
NEW_CHUNKS_PATH = PROJECT_ROOT / "08_chunks" / "chunks_with_major_accident_docs.jsonl"
ORIGINAL_VECTOR_DIR = PROJECT_ROOT / "10_vector_db"
NEW_VECTOR_DIR = PROJECT_ROOT / "10_vector_db_with_major_accident_docs"
TEST_DIR = PROJECT_ROOT / "13_added_docs_tests"
REPORT_PATH = TEST_DIR / "add_major_accident_docs_report.txt"
SEARCH_TEST_PATH = TEST_DIR / "major_accident_docs_search_test.txt"
EXTRACTED_TEXT_DIR = TEST_DIR / "extracted_texts"
COLLECTION_NAME = "mine_safety_docs"
EMBEDDING_MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
CHUNK_SIZE = 1500
OVERLAP = 250
EXPECTED_PDFS = [
    "중대재해처벌법령_FAQ_중대산업재해.pdf",
    "중대재해처벌법_질의회시집.pdf",
    "중대재해처벌법_해설서.pdf",
    "중대재해처벌법_따라하기_안내서.pdf",
]
TEST_QUERIES = [
    "중대재해처벌법 FAQ",
    "경영책임자 안전보건 확보의무",
    "안전보건관리체계 구축",
    "중대재해처벌법 질의회시",
    "중대산업재해 예방",
    "도급 용역 위탁 안전보건 확보의무",
    "유해위험요인 확인 개선 절차",
]


def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as f:
        for line_no, line in enumerate(f, start=1):
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError as exc:
                raise RuntimeError(f"JSONL parse failed at {path}:{line_no}: {exc}") from exc
    return rows


def write_jsonl(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")


def backup_file(path: Path) -> Path | None:
    if not path.exists():
        return None
    backup_path = path.with_name(f"{path.stem}_backup_before_major_accident_docs_{now_stamp()}{path.suffix}")
    shutil.copy2(path, backup_path)
    return backup_path


def backup_dir(path: Path) -> Path | None:
    if not path.exists():
        return None
    backup_path = path.with_name(f"{path.name}_backup_before_major_accident_docs_{now_stamp()}")
    shutil.copytree(path, backup_path)
    return backup_path


def make_chunks(text: str, chunk_size: int = CHUNK_SIZE, overlap: int = OVERLAP) -> list[str]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = "\n".join(line.strip() for line in text.splitlines())
    text = "\n".join(line for line in text.splitlines() if line)

    chunks: list[str] = []
    start = 0
    text_length = len(text)

    while start < text_length:
        end = start + chunk_size
        chunk = text[start:end].strip()
        if chunk:
            chunks.append(chunk)
        if end >= text_length:
            break
        start = end - overlap
    return chunks


def find_pdf_files() -> list[Path]:
    if not RAW_ADDED_DIR.exists():
        raise FileNotFoundError(f"추가 자료 폴더를 찾을 수 없습니다: {RAW_ADDED_DIR}")
    return sorted(
        path
        for path in RAW_ADDED_DIR.rglob("*.pdf*")
        if path.is_file() and ".pdf" in path.name.lower()
    )


def normalize_pdf_name(name: str) -> str:
    text = name.strip()
    while text.lower().endswith(".pdf"):
        text = text[:-4]
    return f"{text}.pdf"


def extract_pdf_text(pdf_path: Path) -> tuple[str, int]:
    try:
        from pypdf import PdfReader
    except Exception as exc:
        raise RuntimeError(
            "pypdf가 필요합니다. 이 단계는 pypdf가 설치된 Python으로 실행해 주세요."
        ) from exc

    reader = PdfReader(str(pdf_path))
    page_count = len(reader.pages)
    text_parts: list[str] = []
    for page_num, page in enumerate(reader.pages, start=1):
        try:
            page_text = page.extract_text() or ""
        except Exception as exc:
            page_text = f"[PAGE {page_num} TEXT EXTRACTION ERROR: {exc}]"
        text_parts.append(f"\n\n===== PAGE {page_num} =====\n\n")
        text_parts.append(page_text)
    return "".join(text_parts).strip(), page_count


def next_chunk_number(rows: list[dict[str, Any]]) -> int:
    max_no = 0
    for row in rows:
        chunk_id = str(row.get("chunk_id", ""))
        if chunk_id.startswith("chunk_"):
            suffix = chunk_id.split("_", 1)[1]
            if suffix.isdigit():
                max_no = max(max_no, int(suffix))
    return max_no + 1


def build_chunks() -> dict[str, Any]:
    TEST_DIR.mkdir(parents=True, exist_ok=True)
    EXTRACTED_TEXT_DIR.mkdir(parents=True, exist_ok=True)

    if not ORIGINAL_CHUNKS_PATH.exists():
        raise FileNotFoundError(f"기존 chunks 파일을 찾을 수 없습니다: {ORIGINAL_CHUNKS_PATH}")

    original_rows = read_jsonl(ORIGINAL_CHUNKS_PATH)
    backup_chunks = backup_file(ORIGINAL_CHUNKS_PATH)
    pdf_files = find_pdf_files()
    normalized_found_names = {normalize_pdf_name(path.name) for path in pdf_files}
    missing_expected = [name for name in EXPECTED_PDFS if name not in normalized_found_names]

    new_rows: list[dict[str, Any]] = []
    next_no = next_chunk_number(original_rows)
    pdf_summaries: list[dict[str, Any]] = []

    for pdf_path in pdf_files:
        normalized_name = normalize_pdf_name(pdf_path.name)
        normalized_stem = normalized_name[:-4]
        text, page_count = extract_pdf_text(pdf_path)
        txt_path = EXTRACTED_TEXT_DIR / f"{normalized_stem}.txt"
        txt_path.write_text(text, encoding="utf-8")
        chunks = make_chunks(text)
        pdf_summaries.append(
            {
                "file_name": normalized_name,
                "raw_file_name": pdf_path.name,
                "page_count": page_count,
                "char_count": len(text),
                "chunk_count": len(chunks),
            }
        )
        for chunk_index, chunk_text in enumerate(chunks, start=1):
            chunk_id = f"chunk_{next_no:05d}"
            next_no += 1
            new_rows.append(
                {
                    "chunk_id": chunk_id,
                    "source_file": f"{normalized_stem}.txt",
                    "source_path": str(pdf_path.relative_to(PROJECT_ROOT)),
                    "source": normalized_name,
                    "doc_name": normalized_stem,
                    "file_name": normalized_name,
                    "raw_file_name": pdf_path.name,
                    "chunk_index": chunk_index,
                    "char_count": len(chunk_text),
                    "category": "중대재해처벌법_추가자료",
                    "added_stage": "after_Q001_Q110_evaluation",
                    "text": chunk_text,
                }
            )

    combined_rows = original_rows + new_rows
    write_jsonl(NEW_CHUNKS_PATH, combined_rows)

    return {
        "backup_chunks": str(backup_chunks) if backup_chunks else "",
        "pdf_files": [str(path) for path in pdf_files],
        "missing_expected": missing_expected,
        "pdf_summaries": pdf_summaries,
        "original_chunk_count": len(original_rows),
        "added_chunk_count": len(new_rows),
        "new_chunk_count": len(combined_rows),
        "new_chunks_path": str(NEW_CHUNKS_PATH),
    }


def find_text(item: dict[str, Any]) -> str:
    for key in ("text", "chunk_text", "content", "page_content", "document"):
        value = item.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    metadata = item.get("metadata")
    if isinstance(metadata, dict):
        for key in ("text", "chunk_text", "content", "page_content", "document"):
            value = metadata.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()
    return ""


def clean_metadata_value(value: Any) -> str | int | float | bool:
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return json.dumps(value, ensure_ascii=False)


def make_metadata(item: dict[str, Any], index: int) -> dict[str, Any]:
    metadata: dict[str, Any] = {}
    raw_metadata = item.get("metadata")
    if isinstance(raw_metadata, dict):
        for key, value in raw_metadata.items():
            metadata[key] = clean_metadata_value(value)
    for key, value in item.items():
        if key in {"text", "chunk_text", "content", "page_content", "document"}:
            continue
        if key in {"metadata", "embedding", "embeddings"}:
            continue
        metadata[key] = clean_metadata_value(value)
    metadata["chunk_index"] = index
    if "source" not in metadata or not metadata.get("source"):
        metadata["source"] = (
            metadata.get("file_name")
            or metadata.get("doc_name")
            or metadata.get("title")
            or metadata.get("source_file")
            or "출처 정보 없음"
        )
    return metadata


def build_vector_db() -> dict[str, Any]:
    try:
        import chromadb
        from sentence_transformers import SentenceTransformer
    except Exception as exc:
        raise RuntimeError(
            "chromadb와 sentence_transformers가 필요합니다. 이 단계는 프로젝트 .venv Python으로 실행해 주세요."
        ) from exc

    if not NEW_CHUNKS_PATH.exists():
        raise FileNotFoundError(f"새 chunks 파일을 찾을 수 없습니다: {NEW_CHUNKS_PATH}")

    if NEW_VECTOR_DIR.exists():
        shutil.rmtree(NEW_VECTOR_DIR)
    NEW_VECTOR_DIR.mkdir(parents=True, exist_ok=True)

    rows = read_jsonl(NEW_CHUNKS_PATH)
    ids: list[str] = []
    documents: list[str] = []
    metadatas: list[dict[str, Any]] = []
    used_ids: set[str] = set()
    for index, item in enumerate(rows):
        text = find_text(item)
        if not text:
            continue
        chunk_id = str(item.get("chunk_id") or f"chunk_{index + 1:05d}")
        if chunk_id in used_ids:
            chunk_id = f"{chunk_id}_{index + 1:05d}"
        used_ids.add(chunk_id)
        ids.append(chunk_id)
        documents.append(text)
        metadatas.append(make_metadata(item, index))

    model = SentenceTransformer(EMBEDDING_MODEL_NAME)
    client = chromadb.PersistentClient(path=str(NEW_VECTOR_DIR))
    collection = client.get_or_create_collection(
        name=COLLECTION_NAME,
        metadata={
            "description": "광산 안전 지침 및 중대재해처벌법 추가자료 반영 RAG Vector DB",
            "project": "mine_safety_rag",
            "added_stage": "after_Q001_Q110_evaluation",
        },
    )

    batch_size = 64
    for start in range(0, len(documents), batch_size):
        end = min(start + batch_size, len(documents))
        embeddings = model.encode(
            documents[start:end],
            normalize_embeddings=True,
            show_progress_bar=False,
        ).tolist()
        collection.add(
            ids=ids[start:end],
            documents=documents[start:end],
            metadatas=metadatas[start:end],
            embeddings=embeddings,
        )
        print(f"Vector DB 저장: {end}/{len(documents)}")

    return {
        "new_vector_dir": str(NEW_VECTOR_DIR),
        "vector_chunk_count": collection.count(),
    }


def query_vector_db(db_dir: Path, queries: list[str]) -> list[dict[str, Any]]:
    import chromadb
    from sentence_transformers import SentenceTransformer

    model = SentenceTransformer(EMBEDDING_MODEL_NAME)
    client = chromadb.PersistentClient(path=str(db_dir))
    collection = client.get_collection(name=COLLECTION_NAME)
    outputs: list[dict[str, Any]] = []
    for query in queries:
        embedding = model.encode([query], normalize_embeddings=True, show_progress_bar=False).tolist()
        result = collection.query(
            query_embeddings=embedding,
            n_results=5,
            include=["documents", "metadatas", "distances"],
        )
        docs = result.get("documents", [[]])[0]
        metas = result.get("metadatas", [[]])[0]
        distances = result.get("distances", [[]])[0]
        hits = []
        for rank, doc in enumerate(docs, start=1):
            meta = metas[rank - 1] if rank - 1 < len(metas) else {}
            distance = distances[rank - 1] if rank - 1 < len(distances) else None
            hits.append(
                {
                    "rank": rank,
                    "source": meta.get("source") or meta.get("file_name") or meta.get("source_file") or "",
                    "chunk_id": meta.get("chunk_id") or "",
                    "category": meta.get("category") or "",
                    "distance": distance,
                    "preview": doc[:240].replace("\n", " "),
                    "is_added_doc": meta.get("added_stage") == "after_Q001_Q110_evaluation",
                }
            )
        outputs.append({"query": query, "hits": hits})
    return outputs


def write_search_test(results: list[dict[str, Any]], path: Path) -> None:
    lines = [
        "중대재해처벌법 추가자료 검색 테스트",
        f"생성 시간: {datetime.now().isoformat(timespec='seconds')}",
        f"Vector DB: {NEW_VECTOR_DIR}",
        "",
    ]
    for item in results:
        lines.append(f"[질문] {item['query']}")
        for hit in item["hits"]:
            marker = "추가자료" if hit["is_added_doc"] else "기존자료"
            distance = hit["distance"]
            distance_text = f"{distance:.4f}" if isinstance(distance, (int, float)) else ""
            lines.append(
                f"  {hit['rank']}. [{marker}] {hit['source']} | {hit['chunk_id']} | "
                f"category={hit['category']} | distance={distance_text}"
            )
            lines.append(f"     {hit['preview']}")
        lines.append("")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


def swap_app_vector_db() -> dict[str, Any]:
    if not NEW_VECTOR_DIR.exists():
        raise FileNotFoundError(f"새 Vector DB 폴더를 찾을 수 없습니다: {NEW_VECTOR_DIR}")
    backup_vector_dir = backup_dir(ORIGINAL_VECTOR_DIR)
    temp_replace_dir = ORIGINAL_VECTOR_DIR.with_name(f"{ORIGINAL_VECTOR_DIR.name}_replace_tmp_{now_stamp()}")
    if temp_replace_dir.exists():
        shutil.rmtree(temp_replace_dir)
    shutil.copytree(NEW_VECTOR_DIR, temp_replace_dir)
    if ORIGINAL_VECTOR_DIR.exists():
        shutil.rmtree(ORIGINAL_VECTOR_DIR)
    temp_replace_dir.rename(ORIGINAL_VECTOR_DIR)
    return {
        "backup_vector_dir": str(backup_vector_dir) if backup_vector_dir else "",
        "app_vector_db_dir": str(ORIGINAL_VECTOR_DIR),
    }


def write_report(
    chunk_info: dict[str, Any] | None,
    vector_info: dict[str, Any] | None,
    swap_info: dict[str, Any] | None,
    search_results: list[dict[str, Any]] | None,
) -> None:
    original_count = chunk_info.get("original_chunk_count") if chunk_info else "확인 필요"
    new_count = chunk_info.get("new_chunk_count") if chunk_info else "확인 필요"
    added_count = chunk_info.get("added_chunk_count") if chunk_info else "확인 필요"
    pdf_files = chunk_info.get("pdf_files", []) if chunk_info else []
    missing_expected = chunk_info.get("missing_expected", []) if chunk_info else []
    pdf_summaries = chunk_info.get("pdf_summaries", []) if chunk_info else []
    search_text = SEARCH_TEST_PATH.read_text(encoding="utf-8") if SEARCH_TEST_PATH.exists() else "검색 테스트 미실행"
    app_path = PROJECT_ROOT / "app.py"
    app_uses_new_db = False
    if app_path.exists():
        try:
            app_uses_new_db = "10_vector_db_with_major_accident_docs" in app_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            app_uses_new_db = False
    app_modified_text = (
        "수정함 - VECTOR_DB_DIR 경로만 10_vector_db_with_major_accident_docs로 변경"
        if app_uses_new_db
        else "수정하지 않음"
    )
    vector_backup_candidates = sorted(
        PROJECT_ROOT.glob("10_vector_db_backup_before_major_accident_docs_*"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    detected_vector_backup = str(vector_backup_candidates[0]) if vector_backup_candidates else ""

    lines = [
        "중대재해처벌법 추가 공식 PDF 자료 RAG 반영 보고서",
        f"생성 시간: {datetime.now().isoformat(timespec='seconds')}",
        "",
        "1. 추가 처리된 PDF 목록",
    ]
    if pdf_files:
        for pdf in pdf_files:
            lines.append(f"- {pdf}")
    else:
        lines.append("- 처리된 PDF 없음")
    if missing_expected:
        lines.append("")
        lines.append("예상 파일 중 현재 폴더에 없는 파일")
        for name in missing_expected:
            lines.append(f"- {name}")

    lines.extend(
        [
            "",
            "PDF별 처리 요약",
        ]
    )
    for item in pdf_summaries:
        lines.append(
            f"- {item['file_name']}: pages={item['page_count']}, chars={item['char_count']}, chunks={item['chunk_count']}"
        )

    lines.extend(
        [
            "",
            f"2. 추가 전 chunk 수: {original_count}",
            f"3. 추가 후 chunk 수: {new_count}",
            f"4. 증가한 chunk 수: {added_count}",
            f"5. 새 chunks 파일 경로: {NEW_CHUNKS_PATH}",
            f"6. 새 Vector DB 폴더 경로: {NEW_VECTOR_DIR}",
            f"7. app.py 수정 여부: {app_modified_text}",
            "8. 기존 평가 파일 수정 여부: 수정하지 않음",
            "9. 기존 비교 실험 파일 수정 여부: 수정하지 않음",
            "10. Q001~Q110 자동평가 재실행 여부: 재실행하지 않음",
            "",
            "11. 검색 테스트 결과",
            search_text,
            "",
            "12. 앞으로 앱 답변에 새 자료가 반영되는지 여부",
        ]
    )
    if swap_info:
        lines.append(
            "- 반영됨. app.py는 기존처럼 10_vector_db를 읽고, 해당 폴더가 새 자료 포함 DB로 교체되었습니다."
        )
    elif app_uses_new_db:
        lines.append(
            "- 반영됨. app.py의 VECTOR_DB_DIR이 10_vector_db_with_major_accident_docs를 읽도록 최소 수정되었습니다."
        )
    else:
        lines.append(
            "- 별도 DB 생성까지 완료됨. app.py 반영은 10_vector_db 교체 또는 경로 변경 후 적용됩니다."
        )

    lines.extend(
        [
            "",
            "13. 롤백 방법",
            "- app.py를 되돌리려면 app_backup_major_accident_docs_*.py 백업 중 최신 파일을 app.py로 복사하거나, VECTOR_DB_DIR 한 줄을 ROOT_DIR / \"10_vector_db\"로 되돌립니다.",
            "- 새 chunks 파일만 제거하려면 08_chunks/chunks_with_major_accident_docs.jsonl 파일을 사용하지 않으면 됩니다.",
            "- 앱 DB를 원래 상태로 되돌리려면 10_vector_db를 삭제하지 말고 이름을 바꾼 뒤, 아래 백업 폴더를 10_vector_db로 복사합니다.",
            f"- Vector DB 백업 폴더: {(swap_info or {}).get('backup_vector_dir') or detected_vector_backup or '아직 교체 전 또는 백업 없음'}",
            f"- chunks.jsonl 백업 파일: {(chunk_info or {}).get('backup_chunks', '백업 없음')}",
            "",
            "보존 확인",
            "- 09_answer_tests 기존 평가 파일: 수정하지 않음",
            "- 12_compare_experiment 기존 비교 실험 결과: 수정하지 않음",
            "- Q001~Q110 평가 점수 및 94.90점 비교 결과: 재계산하지 않음",
        ]
    )
    TEST_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")


def load_json(path: Path) -> dict[str, Any] | None:
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def save_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--build-chunks", action="store_true", help="PDF 추출 및 새 chunks 파일 생성")
    parser.add_argument("--build-vector", action="store_true", help="새 chunks 파일로 별도 Vector DB 생성")
    parser.add_argument("--test-search", action="store_true", help="새 Vector DB 검색 테스트")
    parser.add_argument("--swap-app-db", action="store_true", help="10_vector_db 백업 후 새 DB로 교체")
    parser.add_argument("--write-report", action="store_true", help="최종 보고서 생성")
    args = parser.parse_args()

    TEST_DIR.mkdir(parents=True, exist_ok=True)
    state_path = TEST_DIR / "add_major_accident_docs_state.json"
    state = load_json(state_path) or {}

    if args.build_chunks:
        state["chunk_info"] = build_chunks()
        save_json(state_path, state)
        print(json.dumps(state["chunk_info"], ensure_ascii=False, indent=2))

    if args.build_vector:
        state["vector_info"] = build_vector_db()
        save_json(state_path, state)
        print(json.dumps(state["vector_info"], ensure_ascii=False, indent=2))

    if args.test_search:
        search_results = query_vector_db(NEW_VECTOR_DIR, TEST_QUERIES)
        write_search_test(search_results, SEARCH_TEST_PATH)
        state["search_test_path"] = str(SEARCH_TEST_PATH)
        state["search_summary"] = [
            {
                "query": item["query"],
                "added_doc_hits": sum(1 for hit in item["hits"] if hit["is_added_doc"]),
                "top_source": item["hits"][0]["source"] if item["hits"] else "",
            }
            for item in search_results
        ]
        save_json(state_path, state)
        print(json.dumps(state["search_summary"], ensure_ascii=False, indent=2))

    if args.swap_app_db:
        state["swap_info"] = swap_app_vector_db()
        save_json(state_path, state)
        print(json.dumps(state["swap_info"], ensure_ascii=False, indent=2))

    if args.write_report:
        write_report(
            state.get("chunk_info"),
            state.get("vector_info"),
            state.get("swap_info"),
            None,
        )
        print(f"보고서 저장: {REPORT_PATH}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        sys.exit(1)
