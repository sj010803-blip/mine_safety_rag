from pathlib import Path
from pypdf import PdfReader
import csv

BASE_DIR = Path(__file__).resolve().parents[1]

DOCS_DIR = BASE_DIR / "04_documents"
OUTPUT_DIR = BASE_DIR / "06_processed_texts"
METADATA_DIR = BASE_DIR / "05_metadata"
OUTPUT_CSV = METADATA_DIR / "text_extract_result.csv"

OUTPUT_DIR.mkdir(exist_ok=True)
METADATA_DIR.mkdir(exist_ok=True)

pdf_files = sorted(DOCS_DIR.rglob("*.pdf"))

rows = []

print("=" * 60)
print("Mine Safety RAG PDF text extraction start")
print("=" * 60)
print("DOCS_DIR:", DOCS_DIR)
print("OUTPUT_DIR:", OUTPUT_DIR)
print("PDF count:", len(pdf_files))
print()

for idx, pdf_path in enumerate(pdf_files, start=1):
    relative_path = pdf_path.relative_to(BASE_DIR)
    output_name = pdf_path.stem + ".txt"
    output_path = OUTPUT_DIR / output_name

    status = "OK"
    note = ""
    page_count = 0
    char_count = 0

    print(f"[{idx}/{len(pdf_files)}] {relative_path}")

    try:
        reader = PdfReader(str(pdf_path))
        page_count = len(reader.pages)

        text_parts = []
        for page_num, page in enumerate(reader.pages, start=1):
            try:
                page_text = page.extract_text() or ""
            except Exception as e:
                page_text = ""
                text_parts.append(f"\n\n[PAGE {page_num} TEXT EXTRACTION ERROR: {e}]\n\n")

            text_parts.append(f"\n\n===== PAGE {page_num} =====\n\n")
            text_parts.append(page_text)

        full_text = "".join(text_parts).strip()
        char_count = len(full_text)

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(full_text)

        if char_count < 100:
            status = "CHECK"
            note = "extracted text is very short; PDF may be image-based or protected"

        print(f"   -> pages: {page_count}, chars: {char_count}, status: {status}")

    except Exception as e:
        status = "ERROR"
        note = str(e)
        print("   -> ERROR:", e)

    rows.append({
        "pdf_file": pdf_path.name,
        "pdf_relative_path": str(relative_path),
        "txt_file": output_name,
        "page_count": page_count,
        "char_count": char_count,
        "status": status,
        "note": note
    })

with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(
        f,
        fieldnames=[
            "pdf_file",
            "pdf_relative_path",
            "txt_file",
            "page_count",
            "char_count",
            "status",
            "note"
        ]
    )
    writer.writeheader()
    writer.writerows(rows)

print()
print("=" * 60)
print("PDF text extraction finished")
print("Saved text files:", OUTPUT_DIR)
print("Saved CSV:", OUTPUT_CSV)
print("=" * 60)
