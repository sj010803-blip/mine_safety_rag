from pathlib import Path
import csv

BASE_DIR = Path(__file__).resolve().parents[1]

DOCS_DIR = BASE_DIR / "04_documents"
METADATA_DIR = BASE_DIR / "05_metadata"
OUTPUT_CSV = METADATA_DIR / "document_check_result.csv"

METADATA_DIR.mkdir(exist_ok=True)

pdf_files = list(DOCS_DIR.rglob("*.pdf"))

rows = []

print("=" * 60)
print("Mine Safety RAG document check start")
print("=" * 60)
print("BASE_DIR:", BASE_DIR)
print("DOCS_DIR:", DOCS_DIR)
print("PDF count:", len(pdf_files))
print()

for pdf in sorted(pdf_files):
    relative_path = pdf.relative_to(BASE_DIR)
    file_name = pdf.name

    status = "OK"
    note = ""

    if "구버전" in file_name or "VERSION_CHECK" in file_name:
        status = "CHECK"
        note = "old version or version check file"

    if "09_" in file_name and "제10차개정" not in file_name:
        status = "CHECK"
        note = "09 mine safety technical standard may not be latest"

    if "10_" in file_name and "2025-35" not in file_name:
        status = "CHECK"
        note = "10 mine safety work guideline may not be latest"

    rows.append({
        "file_name": file_name,
        "relative_path": str(relative_path),
        "status": status,
        "note": note
    })

    print(f"[{status}] {relative_path}")
    if note:
        print("   ->", note)

with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(
        f,
        fieldnames=["file_name", "relative_path", "status", "note"]
    )
    writer.writeheader()
    writer.writerows(rows)

print()
print("=" * 60)
print("Document check finished")
print("Saved CSV:", OUTPUT_CSV)
print("=" * 60)