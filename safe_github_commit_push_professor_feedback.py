from __future__ import annotations

import ast
import datetime as dt
import subprocess
import sys
from pathlib import Path


ROOT = Path(r"C:\Users\USER\Desktop\mine_safety_rag")
REPORT = ROOT / "20_eval_visualization" / "git_commit_push_professor_feedback_report.txt"
MAIN_COMMIT_MESSAGE = "feat: apply professor feedback updates"
REPORT_COMMIT_MESSAGE = "docs: add professor feedback git report"

GIT_CANDIDATES = [
    Path(r"C:\Program Files\Git\cmd\git.exe"),
    Path(r"C:\Program Files\Git\bin\git.exe"),
    Path(r"C:\Users\USER\AppData\Local\GitHubDesktop\app-3.6.2\resources\app\git\cmd\git.exe"),
    Path(r"C:\Users\USER\AppData\Local\Programs\Git\cmd\git.exe"),
]

WHITELIST = [
    "app.py",
    "safe_github_commit_push_professor_feedback.py",
    "app_backup_before_professor_feedback_ui.py",
    "app_backup_before_risk_badge_similarity_note.py",
    "app_backup_before_legal_warning_ui_update.py",
    "19_similarity_comparison/similarity_comparison_template.xlsx",
    "19_similarity_comparison/similarity_comparison_summary.xlsx",
    "19_similarity_comparison/similarity_comparison_report.txt",
    "19_similarity_comparison/similarity_comparison_bar_chart.png",
    "20_eval_visualization/model_score_boxplot.png",
    "20_eval_visualization/model_score_mean_std_bar.png",
    "20_eval_visualization/model_score_summary_table.xlsx",
    "20_eval_visualization/eval_visualization_report.txt",
    "20_eval_visualization/professor_feedback_update_report.txt",
    "20_eval_visualization/risk_badge_similarity_note_update_report.txt",
    "20_eval_visualization/legal_warning_ui_update_report.txt",
    # Recent webapp feature improvements explicitly allowed by the request.
    "app_backup_before_legal_evidence_history_features.py",
    "app_backup_before_excel_download_fix.py",
    "app_backup_before_history_ui_polish.py",
    "data/legal_checklist_status.json",
    "data/conversation_history.jsonl",
    "18_legal_evidence_features/legal_evidence_history_feature_report.txt",
    "18_legal_evidence_features/excel_download_button_fix_report.txt",
    "18_legal_evidence_features/history_ui_polish_report.txt",
    "18_legal_evidence_features/legal_checklist_export.xlsx",
    "18_legal_evidence_features/risk_assessment_draft_export.xlsx",
    "18_legal_evidence_features/conversation_history_export.xlsx",
    "17_similarity_index/similarity_summary.xlsx",
    "17_similarity_index/similarity_overall_report.txt",
    "17_similarity_index/similarity_index_bar_chart.png",
    "16_final_outputs/01_final_excel/MineSafe_AI_v2_최종평가_통합보고서_SAFE_WITH_SIMILARITY.xlsx",
    "16_final_outputs/03_presentation/MineSafe_AI_주간발표자료_WITH_SIMILARITY.pptx",
    "16_final_outputs/04_report_text/최종_1페이지_요약_WITH_SIMILARITY.txt",
    "16_final_outputs/06_professor_QA/교수님_예상질문_답변_WITH_SIMILARITY.txt",
]

REPORT_WHITELIST = ["20_eval_visualization/git_commit_push_professor_feedback_report.txt"]

FORBIDDEN_SNIPPETS = [
    ".env",
    ".streamlit/secrets.toml",
    "08_chunks/",
    "10_vector_db/",
    "10_vector_db_with_major_accident_docs/",
    "__pycache__/",
    "~$",
]


def normalize_rel(path: str | Path) -> str:
    return str(path).replace("\\", "/").lstrip("./")


def find_git() -> Path:
    for candidate in GIT_CANDIDATES:
        if candidate.exists():
            return candidate
    raise RuntimeError("git.exe를 찾을 수 없습니다.")


GIT = find_git()


def run_git(args: list[str], check: bool = True) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        [str(GIT), "-c", "core.quotepath=false", *args],
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


def is_forbidden(path: str | Path) -> bool:
    rel = normalize_rel(path)
    lower = rel.lower()
    if lower.endswith(".pyc"):
        return True
    return any(snippet.lower() in lower for snippet in FORBIDDEN_SNIPPETS)


def ast_check_app() -> str:
    app_path = ROOT / "app.py"
    ast.parse(app_path.read_text(encoding="utf-8"))
    return "통과(ast.parse)"


def ensure_safe_candidates(candidates: list[str]) -> tuple[list[str], list[str], list[str]]:
    existing: list[str] = []
    missing: list[str] = []
    forbidden: list[str] = []
    for rel in candidates:
        normalized = normalize_rel(rel)
        if is_forbidden(normalized):
            forbidden.append(normalized)
            continue
        if (ROOT / normalized).exists():
            existing.append(normalized)
        else:
            missing.append(normalized)
    return existing, missing, forbidden


def get_staged() -> list[str]:
    output = run_git(["diff", "--cached", "--name-only"], check=True).stdout
    return [normalize_rel(line.strip()) for line in output.splitlines() if line.strip()]


def reset_stage() -> None:
    run_git(["reset"], check=False)


def abort_if_forbidden_staged(staged: list[str], allowed: set[str]) -> None:
    forbidden_staged = [path for path in staged if is_forbidden(path) or path not in allowed]
    if forbidden_staged:
        reset_stage()
        raise RuntimeError(
            "금지 또는 whitelist 밖 파일이 stage되어 stage를 비우고 중단합니다: "
            + ", ".join(forbidden_staged)
        )


def git_add_files(files: list[str]) -> None:
    for rel in files:
        run_git(["add", "-f", "--", rel], check=True)


def write_report(text: str) -> None:
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text(text, encoding="utf-8")


def main() -> int:
    started = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report_lines: list[str] = []
    main_commit_hash = ""
    report_commit_hash = ""
    push_success = False
    report_push_success = False
    origin_match = False

    try:
        if Path.cwd().resolve() != ROOT.resolve():
            raise RuntimeError(f"현재 경로가 프로젝트 경로가 아닙니다: {Path.cwd()}")

        status_before = run_git(["status", "--short"], check=True).stdout
        branch = run_git(["rev-parse", "--abbrev-ref", "HEAD"], check=True).stdout.strip()
        if branch != "main":
            raise RuntimeError(f"현재 브랜치가 main이 아닙니다: {branch}")

        app_syntax_result = ast_check_app()
        existing, missing, forbidden_candidates = ensure_safe_candidates(WHITELIST)

        if forbidden_candidates:
            raise RuntimeError("whitelist 후보에 금지 패턴 파일이 있습니다: " + ", ".join(forbidden_candidates))

        pre_staged = get_staged()
        abort_if_forbidden_staged(pre_staged, set(existing))

        git_add_files(existing)
        staged = get_staged()
        abort_if_forbidden_staged(staged, set(existing))

        if not staged:
            raise RuntimeError("stage된 파일이 없어 커밋을 생성하지 않았습니다.")

        print("=== STAGED FILES BEFORE COMMIT ===")
        for path in staged:
            print(path)

        commit_result = run_git(["commit", "-m", MAIN_COMMIT_MESSAGE], check=True)
        main_commit_hash = run_git(["rev-parse", "HEAD"], check=True).stdout.strip()
        push_result = run_git(["push", "origin", "main"], check=True)
        push_success = True

        status_after_push = run_git(["status", "--short"], check=True).stdout
        local_head = run_git(["rev-parse", "HEAD"], check=True).stdout.strip()
        remote_head = run_git(["rev-parse", "origin/main"], check=False).stdout.strip()
        origin_match = local_head == remote_head

        forbidden_stage_check = "통과"
        uncommitted_report_note = "보고서 파일은 main 커밋 후 생성되어 별도 작은 커밋으로 추가"

        report_lines = [
            "MineSafe AI 교수님 피드백 반영 Git commit/push 보고서",
            "",
            f"- 검사 시간: {started}",
            f"- 프로젝트 경로: {ROOT}",
            f"- 사용 Git: {GIT}",
            f"- app.py 문법 검사 결과: {app_syntax_result}",
            f"- whitelist 후보 파일 수: {len(WHITELIST)}",
            f"- 실제 존재한 파일 수: {len(existing)}",
            f"- 누락된 whitelist 후보 수: {len(missing)}",
            "",
            "1. 실제 stage된 파일 목록",
            *[f"- {path}" for path in staged],
            "",
            "2. 누락된 whitelist 후보 파일",
            *([f"- {path}" for path in missing] if missing else ["- 없음"]),
            "",
            "3. 금지 패턴 검사 결과",
            f"- stage 금지 파일 검사: {forbidden_stage_check}",
            "- .env/API Key/Vector DB/chunks/__pycache__/*.pyc/~$* whitelist 제외 확인",
            "",
            "4. 커밋 및 push",
            f"- 메인 커밋 메시지: {MAIN_COMMIT_MESSAGE}",
            f"- 메인 커밋 해시: {main_commit_hash}",
            f"- push 성공 여부: {'성공' if push_success else '실패'}",
            f"- origin/main 일치 여부(메인 커밋 기준): {'일치' if origin_match else '불일치 또는 확인 필요'}",
            f"- 보고서 처리: {uncommitted_report_note}",
            "",
            "5. git status 결과(메인 push 직후)",
            status_after_push if status_after_push.strip() else "clean",
            "",
            "6. 커밋 출력",
            commit_result.stdout.strip(),
            "",
            "7. push 출력",
            push_result.stdout.strip(),
        ]
        write_report("\n".join(report_lines) + "\n")

        # Add the report in a second small whitelist-only commit.
        report_existing, report_missing, report_forbidden = ensure_safe_candidates(REPORT_WHITELIST)
        if report_forbidden or report_missing or report_existing != REPORT_WHITELIST:
            raise RuntimeError("보고서 파일 whitelist 검사 실패")
        git_add_files(report_existing)
        report_staged = get_staged()
        abort_if_forbidden_staged(report_staged, set(REPORT_WHITELIST))
        print("=== STAGED REPORT FILE BEFORE SECOND COMMIT ===")
        for path in report_staged:
            print(path)
        run_git(["commit", "-m", REPORT_COMMIT_MESSAGE], check=True)
        report_commit_hash = run_git(["rev-parse", "HEAD"], check=True).stdout.strip()
        run_git(["push", "origin", "main"], check=True)
        report_push_success = True

        final_status = run_git(["status", "--short"], check=True).stdout
        final_local = run_git(["rev-parse", "HEAD"], check=True).stdout.strip()
        final_remote = run_git(["rev-parse", "origin/main"], check=False).stdout.strip()
        final_match = final_local == final_remote

        print("=== FINAL RESULT ===")
        print(f"MAIN_COMMIT={main_commit_hash}")
        print(f"REPORT_COMMIT={report_commit_hash}")
        print(f"FINAL_ORIGIN_MATCH={final_match}")
        print("FINAL_STATUS:")
        print(final_status if final_status.strip() else "clean")
        return 0

    except Exception as exc:
        try:
            staged_now = get_staged()
            if any(is_forbidden(path) for path in staged_now):
                reset_stage()
        except Exception:
            staged_now = []
        error_report = [
            "MineSafe AI 교수님 피드백 반영 Git commit/push 보고서",
            "",
            f"- 검사 시간: {started}",
            f"- 프로젝트 경로: {ROOT}",
            "- 결과: 실패",
            f"- 오류: {exc}",
            f"- app.py 문법 검사 결과: {'수행 전 실패 또는 오류' if 'app_syntax_result' not in locals() else app_syntax_result}",
            f"- 현재 staged 파일: {', '.join(staged_now) if staged_now else '없음'}",
            "- .env/API Key/Vector DB/chunks/__pycache__/*.pyc/~$* 커밋 방지 검사 적용",
        ]
        write_report("\n".join(error_report) + "\n")
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
