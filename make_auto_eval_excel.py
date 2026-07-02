import re
from pathlib import Path

import pandas as pd

base = Path("09_answer_tests")

file1 = base / "auto_eval_Q031_Q065.tsv"
file2 = base / "auto_eval_Q066_Q100.tsv"

if not file1.exists():
    raise FileNotFoundError(f"파일이 없습니다: {file1}")

if not file2.exists():
    raise FileNotFoundError(f"파일이 없습니다: {file2}")

df1 = pd.read_csv(file1, sep="\t")
df2 = pd.read_csv(file2, sep="\t")

df = pd.concat([df1, df2], ignore_index=True)

for column in df.columns:
    if df[column].dtype == "object":
        df[column] = df[column].astype(str).apply(
            lambda value: re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(value))
        )
        df[column] = df[column].replace({"\u0000": ""})

safe_columns = [
    "question_id",
    "category",
    "question",
    "expected_keywords",
    "retrieved_chunk_ids",
    "retrieved_sources",
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

available_columns = [col for col in safe_columns if col in df.columns]
df = df[available_columns].copy()

out = base / "auto_eval_70_detail_report_readable.xlsx"

low = df.sort_values("총점").head(20)

grade = df["판정"].value_counts().reset_index()
grade.columns = ["판정", "문항수"]

cat = (
    df.groupby("category", dropna=False)
    .agg(
        문항수=("question_id", "count"),
        평균점수=("총점", "mean"),
        검토필요수=("검토필요", lambda x: (x.astype(str) == "Y").sum()),
    )
    .reset_index()
)

cat["평균점수"] = cat["평균점수"].round(2)

with pd.ExcelWriter(out, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="자동평가_상세70문항", index=False)
    low.to_excel(writer, sheet_name="낮은점수_TOP20", index=False)
    grade.to_excel(writer, sheet_name="판정별_요약", index=False)
    cat.to_excel(writer, sheet_name="카테고리별_요약", index=False)

print("엑셀 저장 완료:", out)
print("전체 문항 수:", len(df))
print("평균 점수:", round(df["총점"].mean(), 2))
print("검토필요 Y 개수:", (df["검토필요"].astype(str) == "Y").sum())
