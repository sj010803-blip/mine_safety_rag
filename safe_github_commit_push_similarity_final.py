from __future__ import annotations

import re
import subprocess
from datetime import datetime
from pathlib import Path


ROOT = Path(r"C:\Users\USER\Desktop\mine_safety_rag")
GIT = Path(r"C:\Users\USER\AppData\Local\GitHubDesktop\app-3.6.2\resources\app\git\cmd\git.exe")
REPORT = ROOT / "19_similarity_comparison" / "git_commit_push_similarity_final_report.txt"
COMMIT_MESSAGE = "eval: add final similarity comparison results"
REPORT_COMMIT_MESSAGE = "eval: add similarity comparison git report"

WHITELIST = [
    "19_similarity_comparison/similarity_answer_collection_template.xlsx",
    "19_similarity_comparison/similarity_answer_collection_filled_from_raw.xlsx",
    "19_similarity_comparison/similarity_answer_collection_template_backup_before_raw_import.xlsx",
    "19_similarity_comparison/calculate_similarity_comparison.py",
    "19_similarity_comparison/similarity_comparison_pairwise_scores.xlsx",
    "19_similarity_comparison/similarity_comparison_model_summary.xlsx",
    "19_similarity_comparison/similarity_comparison_question_summary.xlsx",
    "19_similarity_comparison/similarity_comparison_final_report.txt",
    "19_similarity_comparison/similarity_comparison_final_bar_chart.png",
    "19_similarity_comparison/similarity_actual_eval_setup_report.txt",
    "19_similarity_comparison/raw_answer_import_and_similarity_report.txt",
    "19_similarity_comparison/raw_answer_reimport_final_similarity_report.txt",
    "19_similarity_comparison/raw_answers/ChatGPT_repeat_1.txt",
    "19_similarity_comparison/raw_answers/ChatGPT_repeat_2.txt",
    "19_similarity_comparison/raw_answers/ChatGPT_repeat_3.txt",
    "19_similarity_comparison/raw_answers/ChatGPT_repeat_4.txt",
    "19_similarity_comparison/raw_answers/ChatGPT_repeat_5.txt",
    "19_similarity_comparison/raw_answers/Gemini_repeat_1.txt",
    "19_similarity_comparison/raw_answers/Gemini_repeat_2.txt",
    "19_similarity_comparison/raw_answers/Gemini_repeat_3.txt",
    "19_similarity_comparison/raw_answers/Gemini_repeat_4.txt",
    "19_similarity_comparison/raw_answers/Gemini_repeat_5.txt",
]

FORBIDDEN_PATH_PARTS = [
    ".env",
    ".streamlit/secrets.toml",
    "08_chunks/",
    "10_vector_db/",
    "10_vector_db_with_major_accident_docs/",
    "__pycache__",
]
FORBIDDEN_SUFFIXES = [".pyc"]
FORBIDDEN_BASENAME_PREFIXES = ["~$"]

SECRET_REGEXES = [
    re.compile(r"sk-[A-Za-z0-9_-]{20,}"),
    re.compile(r"AIza[0-9A-Za-z_-]{20,}"),
    re.compile(r"ghp_[0-9A-Za-z]{20,}"),
    re.compile(r"xox[baprs]-[0-9A-Za-z-]{20,}"),
    re.compile(r"(?i)\b(api[_-]?key|secret[_-]?key|password|token)\s*[:=]\s*['\"]?[A-Za-z0-9_./+=-]{12,}"),
]

FINAL_SCORES = {
    "ChatGPT": {"mean": "41.73", "std": "7.98", "status": "COMPLETE"},
    "Gemini": {"mean": "40.09", "std": "9.57", "status": "COMPLETE"},
    "MineSafe AI": {"mean": "100.00", "std": "0.00", "status": "COMPLETE"},
}


def rel(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def run_git(args: list[str], check: bool = True) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        [str(GIT), *args],
        cwd=str(ROOT),
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    if check and result.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} 실패\n{result.stdout}")
    return result


def is_forbidden_path(path_text: str) -> str | None:
    normalized = path_text.replace("\\", "/")
    basename = Path(normalized).name
    for prefix in FORBIDDEN_BASENAME_PREFIXES:
        if basename.startswith(prefix):
            return f"임시 파일 접두사 {prefix}"
    for suffix in FORBIDDEN_SUFFIXES:
        if normalized.lower().endswith(suffix):
            return f"금지 확장자 {suffix}"
    for part in FORBIDDEN_PATH_PARTS:
        if part.lower() in normalized.lower():
            return f"금지 경로 패턴 {part}"
    return None


def scan_secret_like_text(path: Path) -> list[str]:
    if path.suffix.lower() not in {".txt", ".py", ".tsv", ".md", ".json", ".csv"}:
        return []
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    hits = []
    for pattern in SECRET_REGEXES:
        if pattern.search(text):
            hits.append(pattern.pattern)
    return hits


def get_staged() -> list[str]:
    result = run_git(["diff", "--cached", "--name-only"], check=True)
    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


def write_report(lines: list[str]) -> None:
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text("\n".join(lines) + "\n", encoding="utf-8")


def verify_no_forbidden_staged(staged: list[str]) -> list[str]:
    problems = []
    allowed = set(WHITELIST + [rel(REPORT)])
    for item in staged:
        reason = is_forbidden_path(item)
        if reason:
            problems.append(f"{item}: {reason}")
        if item not in allowed:
            problems.append(f"{item}: whitelist 외 파일")
    return problems


def main() -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines: list[str] = [
        "3모델 유사성 지수 비교 평가 최종 결과 Git commit/push 보고서",
        "",
        f"- 검사 시간: {now}",
        f"- 프로젝트 경로: {ROOT}",
        f"- Git 실행 파일: {GIT}",
        f"- 커밋 메시지: {COMMIT_MESSAGE}",
        "",
    ]

    if not GIT.exists():
        raise RuntimeError(f"Git 실행 파일을 찾을 수 없습니다: {GIT}")

    initial_status = run_git(["status", "--short"]).stdout.strip()
    initial_staged = get_staged()
    lines.append("1. 초기 git status --short")
    lines.append(initial_status if initial_status else "- clean")
    lines.append("")
    lines.append("2. 초기 staged 파일")
    lines.extend(f"- {item}" for item in initial_staged) if initial_staged else lines.append("- 없음")
    lines.append("")

    if initial_staged:
        problems = verify_no_forbidden_staged(initial_staged)
        if problems:
            run_git(["reset"], check=False)
            lines.append("초기 staged 파일에 문제가 있어 git reset으로 스테이징을 비우고 중단했습니다.")
            lines.extend(f"- {p}" for p in problems)
            write_report(lines)
            raise RuntimeError("초기 staged 파일 안전 검사 실패")

    existing: list[str] = []
    missing: list[str] = []
    blocked: list[str] = []
    raw_included: list[str] = []
    secret_scan_results: list[str] = []

    for item in WHITELIST:
        reason = is_forbidden_path(item)
        path = ROOT / item
        if reason:
            blocked.append(f"{item}: {reason}")
            continue
        if not path.exists():
            missing.append(item)
            continue
        hits = scan_secret_like_text(path)
        if hits:
            blocked.append(f"{item}: 비밀키 유사 패턴 감지")
            secret_scan_results.append(f"{item}: BLOCKED")
            continue
        secret_scan_results.append(f"{item}: OK")
        existing.append(item)
        if item.startswith("19_similarity_comparison/raw_answers/"):
            raw_included.append(item)

    lines.append("3. whitelist 후보 파일 검사")
    lines.append(f"- whitelist 후보 파일 수: {len(WHITELIST)}")
    lines.append(f"- 실제 존재한 파일 수: {len(existing)}")
    lines.append(f"- 누락 파일 수: {len(missing)}")
    lines.append(f"- 차단 파일 수: {len(blocked)}")
    lines.append("")
    lines.append("4. 누락 파일 목록")
    lines.extend(f"- {item}" for item in missing) if missing else lines.append("- 없음")
    lines.append("")
    lines.append("5. 금지 패턴 검사 결과")
    lines.extend(f"- {item}" for item in blocked) if blocked else lines.append("- 금지 경로/비밀키 유사 패턴 감지 없음")
    lines.append("")
    lines.append("6. raw_answers 포함 여부")
    lines.append(f"- 포함 파일 수: {len(raw_included)}/10")
    lines.extend(f"- {item}" for item in raw_included)
    lines.append("")

    if blocked:
        write_report(lines)
        raise RuntimeError("금지 패턴 또는 비밀키 유사 패턴이 감지되어 중단")
    if not existing:
        write_report(lines)
        raise RuntimeError("stage할 whitelist 파일이 없습니다.")

    for item in existing:
        run_git(["add", "-f", "--", item])

    staged = get_staged()
    staged_problems = verify_no_forbidden_staged(staged)
    lines.append("7. commit 전 staged 파일 목록")
    lines.extend(f"- {item}" for item in staged) if staged else lines.append("- 없음")
    lines.append("")

    if staged_problems:
        run_git(["reset"], check=False)
        lines.append("8. staged 안전 검사 실패로 git reset 수행 후 중단")
        lines.extend(f"- {p}" for p in staged_problems)
        write_report(lines)
        raise RuntimeError("staged 안전 검사 실패")

    lines.append("8. staged 안전 검사")
    lines.append("- 통과")
    lines.append("")

    commit_hash = "NO_COMMIT"
    push_success = False
    if staged:
        commit_result = run_git(["commit", "-m", COMMIT_MESSAGE])
        lines.append("9. 첫 번째 commit 결과")
        lines.append(commit_result.stdout.strip())
        commit_hash = run_git(["rev-parse", "HEAD"]).stdout.strip()
        push_result = run_git(["push", "origin", "main"])
        push_success = push_result.returncode == 0
        lines.append("")
        lines.append("10. 첫 번째 push 결과")
        lines.append(push_result.stdout.strip())
    else:
        lines.append("9. 첫 번째 commit 결과")
        lines.append("- staged 파일 없음")

    lines.append("")
    lines.append("11. 모델별 최종 유사성 지수")
    for model, score in FINAL_SCORES.items():
        lines.append(f"- {model}: 평균 {score['mean']}, 표준편차 {score['std']}, 상태 {score['status']}")
    lines.append("")
    lines.append("12. 미커밋 금지 항목 확인")
    lines.append("- .env/API Key/Vector DB/chunks/__pycache__/*.pyc/~$* 미커밋 확인: staged 안전 검사 통과")
    lines.append("- 유사성 지수는 정확도 평가가 아니라 동일 질문 반복 답변의 구조와 핵심 안전조치 일관성을 보는 보조 지표")
    lines.append("")

    # Save report after first push. It is intentionally committed in a second small whitelist commit.
    # Do not rewrite the report after that commit, otherwise the report file would become dirty again.
    lines.extend(
        [
            "13. 보고서 파일 커밋 안내",
            "- 이 보고서는 첫 번째 결과 커밋과 push 결과를 기록한 뒤 별도 whitelist 커밋으로 추가됩니다.",
            "- 보고서 커밋 해시는 self-referential 파일 해시 문제를 피하기 위해 콘솔/최종 응답에서 별도로 보고합니다.",
            "",
        ]
    )
    write_report(lines)

    run_git(["add", "-f", "--", rel(REPORT)])
    report_staged = get_staged()
    report_problems = verify_no_forbidden_staged(report_staged)
    if report_problems:
        run_git(["reset"], check=False)
        raise RuntimeError("보고서 staged 안전 검사 실패")

    second_commit_hash = "NO_COMMIT"
    if report_staged:
        second_commit = run_git(["commit", "-m", REPORT_COMMIT_MESSAGE])
        second_commit_hash = run_git(["rev-parse", "HEAD"]).stdout.strip()
        second_push = run_git(["push", "origin", "main"])
        push_success = push_success and second_push.returncode == 0

    final_status = run_git(["status", "--short"]).stdout.strip()
    local_head = run_git(["rev-parse", "HEAD"]).stdout.strip()
    origin_head = run_git(["rev-parse", "origin/main"]).stdout.strip()
    origin_match = local_head == origin_head

    print("SIMILARITY_FINAL_COMMIT_PUSH_DONE")
    print(f"first_commit={commit_hash}")
    print(f"report_commit={second_commit_hash}")
    print(f"push_success={push_success}")
    print(f"origin_match={origin_match}")
    print(f"report={REPORT}")


if __name__ == "__main__":
    main()
