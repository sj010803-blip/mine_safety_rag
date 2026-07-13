from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit


PROJECT_ROOT = Path(__file__).resolve().parent
REPORT_PATH = (
    PROJECT_ROOT
    / "12_compare_experiment"
    / "git_commit_push_report.txt"
)
COMMIT_MESSAGE = "Finalize mine safety RAG dashboard and evaluation results"
MAX_STAGED_FILES = 200

EXACT_WHITELIST = (
    "app.py",
    "README.md",
    "README_DEPLOY.md",
    "requirements.txt",
    ".gitignore",
    ".gitattributes",
    "02_질문시나리오/question_scenarios_110.tsv",
    "08_chunks/chunks.jsonl",
    "09_answer_tests/auto_eval_Q001_Q110.tsv",
    "09_answer_tests/auto_eval_110_full_summary.tsv",
    "09_answer_tests/auto_eval_110_teacher_view.xlsx",
    "12_compare_experiment/llm_rag_comparison_result_final.xlsx",
    "12_compare_experiment/llm_rag_comparison_result_polished.xlsx",
    "12_compare_experiment/github_share_check_report.txt",
    "check_github_share_ready.py",
    "run_llm_rag_comparison_analysis.py",
    "make_llm_comparison_template.py",
    "_llm_rag_workbook_bridge.mjs",
)

DEFAULT_REQUIREMENTS = "\n".join(
    (
        "streamlit",
        "pandas",
        "numpy",
        "chromadb",
        "google-genai",
        "python-dotenv",
        "openpyxl",
        "",
    )
)


@dataclass
class RunResult:
    git_path: str = ""
    branch: str = ""
    remote_origin: str = ""
    staged_files: list[str] = field(default_factory=list)
    missing_whitelist_files: list[str] = field(default_factory=list)
    forbidden_staged_files: list[str] = field(default_factory=list)
    unexpected_staged_files: list[str] = field(default_factory=list)
    commit_success: bool = False
    push_success: bool = False
    commit_status: str = "실행 전"
    push_status: str = "실행 전"
    final_judgment: str = "확인 필요"


def configure_console() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")


def version_key(path: Path) -> tuple[int, ...]:
    match = re.search(r"app-(\d+(?:\.\d+)*)", str(path), re.IGNORECASE)
    if not match:
        return ()
    return tuple(int(part) for part in match.group(1).split("."))


def running_github_desktop_dirs() -> list[Path]:
    if os.name != "nt":
        return []
    command = [
        "powershell.exe",
        "-NoProfile",
        "-Command",
        (
            "Get-Process GitHubDesktop -ErrorAction SilentlyContinue | "
            "Where-Object { $_.Path } | "
            "ForEach-Object { Split-Path -Parent $_.Path } | "
            "Sort-Object -Unique"
        ),
    ]
    try:
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
        )
    except (OSError, subprocess.SubprocessError):
        return []
    return [
        Path(line.strip())
        for line in completed.stdout.splitlines()
        if line.strip()
    ]


def github_desktop_candidates() -> list[Path]:
    candidates: list[Path] = []
    local_app_data = Path(
        os.environ.get(
            "LOCALAPPDATA",
            str(Path.home() / "AppData" / "Local"),
        )
    )
    desktop_root = local_app_data / "GitHubDesktop"

    app_dirs: list[Path] = []
    try:
        app_dirs.extend(desktop_root.glob("app-*"))
    except OSError:
        pass
    app_dirs.extend(running_github_desktop_dirs())
    app_dirs = sorted(set(app_dirs), key=version_key, reverse=True)

    for app_dir in app_dirs:
        candidates.extend(
            (
                app_dir / "resources" / "app" / "git" / "cmd" / "git.exe",
                (
                    app_dir
                    / "resources"
                    / "app"
                    / "git"
                    / "mingw64"
                    / "bin"
                    / "git.exe"
                ),
            )
        )
    return candidates


def git_is_executable(path: Path) -> bool:
    if not path.is_file():
        return False
    try:
        completed = subprocess.run(
            [str(path), "--version"],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
        )
    except (OSError, subprocess.SubprocessError):
        return False
    return completed.returncode == 0


def find_git() -> Path:
    candidates: list[Path] = []
    path_git = shutil.which("git")
    if path_git:
        candidates.append(Path(path_git))

    candidates.extend(github_desktop_candidates())
    candidates.extend(
        (
            Path(r"C:\Program Files\Git\cmd\git.exe"),
            Path(r"C:\Program Files\Git\bin\git.exe"),
            Path(r"C:\Program Files (x86)\Git\cmd\git.exe"),
        )
    )

    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate).lower()
        if key in seen:
            continue
        seen.add(key)
        if git_is_executable(candidate):
            return candidate
    raise RuntimeError(
        "Git 실행 파일을 찾지 못했습니다. GitHub Desktop을 실행한 뒤 다시 시도하세요."
    )


def run_git(
    git_path: Path,
    arguments: list[str],
    *,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    completed = subprocess.run(
        [
            str(git_path),
            "-c",
            "core.quotepath=false",
            "-C",
            str(PROJECT_ROOT),
            *arguments,
        ],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if check and completed.returncode != 0:
        summary = summarize_error(completed.stderr or completed.stdout)
        raise RuntimeError(
            f"git {' '.join(arguments[:2])} 실패: {summary}"
        )
    return completed


def summarize_error(text: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return "원인을 확인할 수 없습니다."
    summary = lines[-1]
    summary = re.sub(
        r"(https?://)[^/@\s]+@",
        r"\1****@",
        summary,
        flags=re.IGNORECASE,
    )
    return summary[:500]


def sanitize_remote(remote: str) -> str:
    remote = remote.strip()
    try:
        parsed = urlsplit(remote)
    except ValueError:
        return re.sub(
            r"(https?://)[^/@\s]+@",
            r"\1****@",
            remote,
            flags=re.IGNORECASE,
        )
    if parsed.scheme not in {"http", "https"} or "@" not in parsed.netloc:
        return remote
    host = parsed.netloc.rsplit("@", 1)[-1]
    return urlunsplit(
        (parsed.scheme, f"****@{host}", parsed.path, parsed.query, parsed.fragment)
    )


def is_forbidden_path(relative_path: str) -> bool:
    normalized = relative_path.replace("\\", "/").strip("/")
    lowered = normalized.lower()
    parts = lowered.split("/")
    name = parts[-1]

    if lowered in {".env", ".streamlit/secrets.toml", "app copy.py"}:
        return True
    if any(
        part in {
            "__pycache__",
            ".venv",
            ".artifact_tool_runtime",
        }
        for part in parts
    ):
        return True
    if name.startswith("~$"):
        return True
    if "backup" in name or "app_backup" in lowered:
        return True
    if ".bak" in name:
        return True
    if Path(name).suffix.lower() in {
        ".pyc",
        ".pyo",
        ".pyd",
        ".tmp",
    }:
        return True
    return False


def ensure_requirements() -> None:
    requirements_path = PROJECT_ROOT / "requirements.txt"
    if not requirements_path.exists():
        requirements_path.write_text(DEFAULT_REQUIREMENTS, encoding="utf-8")


def collect_whitelist() -> tuple[list[str], list[str]]:
    existing: set[str] = set()
    missing: list[str] = []

    for relative_path in EXACT_WHITELIST:
        path = PROJECT_ROOT / relative_path
        if path.is_file():
            existing.add(path.relative_to(PROJECT_ROOT).as_posix())
        else:
            missing.append(relative_path)

    vector_root = PROJECT_ROOT / "10_vector_db"
    if vector_root.is_dir():
        for path in vector_root.rglob("*"):
            if not path.is_file():
                continue
            relative_path = path.relative_to(PROJECT_ROOT).as_posix()
            if not is_forbidden_path(relative_path):
                existing.add(relative_path)

    safe_existing = sorted(
        path
        for path in existing
        if not is_forbidden_path(path)
    )
    return safe_existing, sorted(missing)


def staged_names(git_path: Path) -> list[str]:
    completed = run_git(
        git_path,
        ["diff", "--cached", "--name-only", "-z"],
    )
    return sorted(
        name
        for name in completed.stdout.split("\0")
        if name
    )


def reset_stage(git_path: Path) -> None:
    completed = run_git(git_path, ["reset"], check=False)
    if completed.returncode != 0:
        raise RuntimeError(
            "기존 stage 해제에 실패했습니다: "
            + summarize_error(completed.stderr or completed.stdout)
        )


def write_report(result: RunResult) -> None:
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "GitHub 안전 커밋·Push 결과 보고서",
        "=" * 72,
        f"실행 시간: {datetime.now().astimezone().strftime('%Y-%m-%d %H:%M:%S %Z')}",
        f"프로젝트 경로: {PROJECT_ROOT}",
        f"Git 실행 파일 경로: {result.git_path or '확인 실패'}",
        f"현재 branch: {result.branch or '확인 실패'}",
        f"remote origin: {result.remote_origin or '확인 실패'}",
        f"stage된 파일 수: {len(result.staged_files)}개",
        f"commit 성공 여부: {'예' if result.commit_success else '아니오'}",
        f"commit 상태: {result.commit_status}",
        f"push 성공 여부: {'예' if result.push_success else '아니오'}",
        f"push 상태: {result.push_status}",
        (
            "금지 파일 stage 여부: "
            f"{'예' if result.forbidden_staged_files else '아니오'}"
        ),
        (
            "whitelist 외 stage 여부: "
            f"{'예' if result.unexpected_staged_files else '아니오'}"
        ),
        "",
        "[stage된 파일 목록]",
        *([f"- {path}" for path in result.staged_files] or ["- 없음"]),
        "",
        "[존재하지 않아 제외된 whitelist 파일]",
        *(
            [f"- {path}" for path in result.missing_whitelist_files]
            or ["- 없음"]
        ),
        "",
        "[금지 stage 파일]",
        *(
            [f"- {path}" for path in result.forbidden_staged_files]
            or ["- 없음"]
        ),
        "",
        "[whitelist 외 stage 파일]",
        *(
            [f"- {path}" for path in result.unexpected_staged_files]
            or ["- 없음"]
        ),
        "",
        f"최종 판단: {result.final_judgment}",
        "",
        "주의: API Key 값은 검사·출력·보고서에 포함하지 않았습니다.",
    ]
    REPORT_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


def dry_run(git_path: Path, result: RunResult) -> int:
    whitelist, missing = collect_whitelist()
    result.missing_whitelist_files = missing
    forbidden = [path for path in whitelist if is_forbidden_path(path)]

    print("[DRY RUN] Git 저장소와 whitelist를 점검했습니다.")
    print(f"Git: {git_path}")
    print(f"branch: {result.branch}")
    print(f"origin: {result.remote_origin}")
    print(f"stage 후보 파일 수: {len(whitelist)}개")
    for path in whitelist:
        print(f"  - {path}")

    if forbidden:
        print("[중단] whitelist 후보에 금지 파일이 포함되어 있습니다.")
        return 1
    if len(whitelist) > MAX_STAGED_FILES:
        print(
            f"[중단] 후보 파일이 {MAX_STAGED_FILES}개를 초과했습니다. "
            "사용자 확인이 필요합니다."
        )
        return 2
    print("[OK] 금지 파일 없음, 200개 이하")
    return 0


def execute(git_path: Path, result: RunResult) -> int:
    ensure_requirements()
    whitelist, missing = collect_whitelist()
    whitelist_set = set(whitelist)
    result.missing_whitelist_files = missing

    if len(whitelist) > MAX_STAGED_FILES:
        result.final_judgment = (
            f"중단: whitelist 후보가 {MAX_STAGED_FILES}개를 초과하여 사용자 확인 필요"
        )
        result.commit_status = "실행 안 함"
        result.push_status = "실행 안 함"
        return 2

    reset_stage(git_path)
    try:
        for relative_path in whitelist:
            run_git(git_path, ["add", "--", relative_path])
    except RuntimeError:
        reset_stage(git_path)
        raise

    result.staged_files = staged_names(git_path)
    result.forbidden_staged_files = [
        path
        for path in result.staged_files
        if is_forbidden_path(path)
    ]
    result.unexpected_staged_files = [
        path
        for path in result.staged_files
        if path not in whitelist_set
    ]

    print("[STAGED FILES]")
    for path in result.staged_files:
        print(f"  - {path}")
    print(f"stage된 파일 수: {len(result.staged_files)}개")

    if (
        result.forbidden_staged_files
        or result.unexpected_staged_files
        or len(result.staged_files) > MAX_STAGED_FILES
    ):
        reset_stage(git_path)
        result.commit_status = "안전 검사 실패로 실행 안 함"
        result.push_status = "실행 안 함"
        result.final_judgment = (
            "중단: 금지 파일, whitelist 외 파일 또는 200개 초과 감지"
        )
        return 3

    if not result.staged_files:
        result.commit_status = "커밋할 변경 없음"
        result.push_status = "실행 안 함"
        result.final_judgment = "안전 검사 완료, 커밋할 변경 없음"
        return 0

    commit = run_git(
        git_path,
        ["commit", "-m", COMMIT_MESSAGE],
        check=False,
    )
    if commit.returncode != 0:
        result.commit_status = summarize_error(
            commit.stderr or commit.stdout
        )
        result.push_status = "commit 실패로 실행 안 함"
        result.final_judgment = "commit 실패, stage 상태 확인 필요"
        return 4

    result.commit_success = True
    result.commit_status = "완료"

    push = run_git(
        git_path,
        ["push", "origin", "main"],
        check=False,
    )
    if push.returncode == 0:
        result.push_success = True
        result.push_status = "완료"
        result.final_judgment = "안전한 whitelist 커밋 및 push 완료"
        print("[OK] origin main push 완료")
        return 0

    result.push_status = summarize_error(push.stderr or push.stdout)
    result.final_judgment = (
        "commit 완료, push 실패: GitHub Desktop에서 Push origin만 누르면 됨"
    )
    print(
        "[안내] push 인증 또는 연결 문제입니다. "
        "GitHub Desktop에서 Push origin만 누르면 됨"
    )
    return 5


def main() -> int:
    configure_console()
    parser = argparse.ArgumentParser(
        description="Whitelist 파일만 안전하게 stage, commit, push합니다."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="stage, commit, push 없이 후보 파일만 검사합니다.",
    )
    arguments = parser.parse_args()
    result = RunResult()

    try:
        git_path = find_git()
        result.git_path = str(git_path)

        inside = run_git(
            git_path,
            ["rev-parse", "--is-inside-work-tree"],
        ).stdout.strip()
        if inside.lower() != "true":
            raise RuntimeError("프로젝트 루트가 Git 저장소가 아닙니다.")

        repository_root = Path(
            run_git(
                git_path,
                ["rev-parse", "--show-toplevel"],
            ).stdout.strip()
        ).resolve()
        if repository_root != PROJECT_ROOT.resolve():
            raise RuntimeError(
                f"저장소 루트가 프로젝트 경로와 다릅니다: {repository_root}"
            )

        result.branch = run_git(
            git_path,
            ["branch", "--show-current"],
        ).stdout.strip()
        if result.branch != "main":
            raise RuntimeError(
                f"현재 branch가 main이 아닙니다: {result.branch or '(detached HEAD)'}"
            )

        raw_remote = run_git(
            git_path,
            ["remote", "get-url", "origin"],
        ).stdout.strip()
        result.remote_origin = sanitize_remote(raw_remote)

        if arguments.dry_run:
            return dry_run(git_path, result)
        return execute(git_path, result)
    except Exception as error:
        result.commit_status = "실패"
        result.push_status = "실행 안 함"
        result.final_judgment = f"중단: {str(error)[:500]}"
        print(f"[중단] {result.final_judgment}")
        return 1
    finally:
        if not arguments.dry_run:
            write_report(result)
            print(f"보고서 저장 경로: {REPORT_PATH}")


if __name__ == "__main__":
    raise SystemExit(main())
