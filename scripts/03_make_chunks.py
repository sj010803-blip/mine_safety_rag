from pathlib import Path
import json
import csv

BASE_DIR = Path(__file__).resolve().parents[1]

TEXT_DIR = BASE_DIR / "06_processed_texts"
CHUNK_DIR = BASE_DIR / "08_chunks"
METADATA_DIR = BASE_DIR / "05_metadata"

CHUNK_JSONL = CHUNK_DIR / "chunks.jsonl"
CHUNK_CSV = METADATA_DIR / "chunk_result.csv"

CHUNK_SIZE = 1500
OVERLAP = 250

CHUNK_DIR.mkdir(exist_ok=True)
METADATA_DIR.mkdir(exist_ok=True)

txt_files = sorted(TEXT_DIR.rglob("*.txt"))

def make_chunks(text, chunk_size=CHUNK_SIZE, overlap=OVERLAP):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = "\n".join(line.strip() for line in text.splitlines())
    text = "\n".join(line for line in text.splitlines() if line)

    chunks = []
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

rows = []
total_chunks = 0

print("=" * 60)
print("Mine Safety RAG chunking start")
print("=" * 60)
print("TEXT_DIR:", TEXT_DIR)
print("CHUNK_DIR:", CHUNK_DIR)
print("TXT count:", len(txt_files))
print("CHUNK_SIZE:", CHUNK_SIZE)
print("OVERLAP:", OVERLAP)
print()

with open(CHUNK_JSONL, "w", encoding="utf-8") as jsonl_file:
    for txt_path in txt_files:
        source_name = txt_path.name
        relative_path = txt_path.relative_to(BASE_DIR)

        try:
            text = txt_path.read_text(encoding="utf-8")
            chunks = make_chunks(text)

            print(f"[OK] {relative_path} -> chunks: {len(chunks)}")

            for i, chunk_text in enumerate(chunks, start=1):
                total_chunks += 1
                chunk_id = f"chunk_{total_chunks:05d}"

                record = {
                    "chunk_id": chunk_id,
                    "source_file": source_name,
                    "source_path": str(relative_path),
                    "chunk_index": i,
                    "char_count": len(chunk_text),
                    "text": chunk_text
                }

                jsonl_file.write(json.dumps(record, ensure_ascii=False) + "\n")

                rows.append({
                    "chunk_id": chunk_id,
                    "source_file": source_name,
                    "chunk_index": i,
                    "char_count": len(chunk_text),
                    "preview": chunk_text[:120].replace("\n", " ")
                })

        except Exception as e:
            print(f"[ERROR] {relative_path} -> {e}")

with open(CHUNK_CSV, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(
        f,
        fieldnames=["chunk_id", "source_file", "chunk_index", "char_count", "preview"]
    )
    writer.writeheader()
    writer.writerows(rows)

print()
print("=" * 60)
print("Chunking finished")
print("Total chunks:", total_chunks)
print("Saved JSONL:", CHUNK_JSONL)
print("Saved CSV:", CHUNK_CSV)
print("=" * 60)
