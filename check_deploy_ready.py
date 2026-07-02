from pathlib import Path

ROOT = Path(__file__).resolve().parent

checks = [
    ("app.py", ROOT / "app.py"),
    ("requirements.txt", ROOT / "requirements.txt"),
    (".gitignore", ROOT / ".gitignore"),
    ("08_chunks/chunks.jsonl", ROOT / "08_chunks" / "chunks.jsonl"),
    ("10_vector_db", ROOT / "10_vector_db"),
    ("02_질문시나리오/question_scenarios_30.tsv", ROOT / "02_질문시나리오" / "question_scenarios_30.tsv"),
    ("02_질문시나리오/question_scenarios_65.tsv", ROOT / "02_질문시나리오" / "question_scenarios_65.tsv"),
    ("02_질문시나리오/question_scenarios_100.tsv", ROOT / "02_질문시나리오" / "question_scenarios_100.tsv"),
    ("09_answer_tests/auto_eval_100_summary.tsv", ROOT / "09_answer_tests" / "auto_eval_100_summary.tsv"),
    ("09_answer_tests/auto_eval_70_detail_report_readable.xlsx", ROOT / "09_answer_tests" / "auto_eval_70_detail_report_readable.xlsx"),
]

print("배포 준비 점검을 시작합니다.")
for label, path in checks:
    if path.exists():
        print(f"[OK] {label} 확인")
    else:
        print(f"[실패] {label} 누락")

env_path = ROOT / ".env"
if env_path.exists():
    gitignore_path = ROOT / ".gitignore"
    if gitignore_path.exists():
        gitignore_text = gitignore_path.read_text(encoding="utf-8")
        if ".env" in gitignore_text:
            print("[OK] .env 파일이 .gitignore에 포함되어 있습니다.")
        else:
            print("[주의] .env 파일은 존재하지만 .gitignore에 포함되어 있지 않습니다.")
    else:
        print("[주의] .gitignore 파일이 없습니다.")
else:
    print("[OK] .env 파일이 없어 보호 대상이 없습니다.")

print("배포 준비 점검 완료")
