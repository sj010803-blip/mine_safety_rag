from __future__ import annotations

import fnmatch
import os
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
REPORT_PATH = (
    PROJECT_ROOT
    / "12_compare_experiment"
    / "github_share_check_report.txt"
)

ENV_PATH = PROJECT_ROOT / ".env"
SECRETS_PATH = PROJECT_ROOT / ".streamlit" / "secrets.toml"

CORE_FILES = (
    "app.py",
    "requirements.txt",
    "README.md",
    ".gitignore",
    "02_질문시나리오/question_scenarios_110.tsv",
    "08_chunks/chunks.jsonl",
    "09_answer_tests/auto_eval_Q001_Q110.tsv",
    "09_answer_tests/auto_eval_110_full_summary.tsv",
    "09_answer_tests/auto_eval_110_teacher_view.xlsx",
    "12_compare_experiment/llm_rag_comparison_result_final.xlsx",
    "12_compare_experiment/llm_rag_comparison_result_polished.xlsx",
)

REQUIRED_IGNORE_PATTERNS = (
    ".env",
    ".streamlit/secrets.toml",
    "__pycache__/",
    "*.pyc",
    "*.pyo",
    "*.pyd",
    ".ipynb_checkpoints/",
    "*.log",
    "app_backup_*.py",
    "evaluation_template_backup_*.tsv",
    "~$*.xlsx",
    "*.tmp",
    ".DS_Store",
    "Thumbs.db",
)

SKIP_DIRECTORIES = {
    ".git",
    ".venv",
    ".artifact_tool_runtime",
    "node_modules",
}


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


@dataclass(frozen=True)
class KeyStatus:
    name: str
    has_value: bool
    masked_value: str


@dataclass(frozen=True)
class SecretFileStatus:
    exists: bool
    keys: tuple[KeyStatus, ...]
    read_error: bool = False

    @property
    def has_valid_key(self) -> bool:
        return any(item.has_value for item in self.keys)


def relative_text(path: Path) -> str:
    return path.relative_to(PROJECT_ROOT).as_posix()


def collect_project_files() -> list[Path]:
    files: list[Path] = []
    for current_root, directory_names, file_names in os.walk(
        PROJECT_ROOT,
        followlinks=False,
    ):
        directory_names[:] = sorted(
            name
            for name in directory_names
            if name not in SKIP_DIRECTORIES
        )
        current_path = Path(current_root)
        for file_name in sorted(file_names):
            path = current_path / file_name
            if path.is_file():
                files.append(path)
    return files


def read_gitignore_patterns() -> list[str]:
    gitignore_path = PROJECT_ROOT / ".gitignore"
    if not gitignore_path.is_file():
        return []
    return [
        line.strip()
        for line in gitignore_path.read_text(encoding="utf-8-sig").splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    ]


def matches_ignore_pattern(relative_path: str, pattern: str) -> bool:
    normalized_path = relative_path.replace("\\", "/")
    normalized_pattern = pattern.strip().replace("\\", "/")
    if not normalized_pattern or normalized_pattern.startswith("!"):
        return False

    directory_pattern = normalized_pattern.endswith("/")
    normalized_pattern = normalized_pattern.rstrip("/")

    if "/" in normalized_pattern:
        if directory_pattern:
            return (
                normalized_path == normalized_pattern
                or normalized_path.startswith(normalized_pattern + "/")
            )
        return fnmatch.fnmatch(normalized_path, normalized_pattern)

    path_parts = normalized_path.split("/")
    if directory_pattern:
        return any(
            fnmatch.fnmatch(part, normalized_pattern)
            for part in path_parts[:-1]
        )
    return fnmatch.fnmatch(path_parts[-1], normalized_pattern)


def is_ignored(relative_path: str, patterns: list[str]) -> bool:
    ignored = False
    for pattern in patterns:
        negated = pattern.startswith("!")
        candidate = pattern[1:] if negated else pattern
        if matches_ignore_pattern(relative_path, candidate):
            ignored = not negated
    return ignored


def is_gemini_key_name(name: str) -> bool:
    normalized = name.strip().upper()
    if normalized in {"GEMINI_API_KEY", "GOOGLE_API_KEY"}:
        return True
    return (
        ("GEMINI" in normalized or "GOOGLE" in normalized)
        and "KEY" in normalized
    )


def normalize_secret_value(value: str) -> str:
    normalized = value.strip().rstrip(",")
    if (
        len(normalized) >= 2
        and normalized[0] == normalized[-1]
        and normalized[0] in {'"', "'"}
    ):
        normalized = normalized[1:-1].strip()
    return normalized


def mask_secret(value: str) -> str:
    if not value:
        return "(비어 있음)"
    if len(value) < 8:
        return "****"
    return f"{value[:4]}****{value[-4:]}"


def inspect_secret_file(path: Path) -> SecretFileStatus:
    if not path.is_file():
        return SecretFileStatus(exists=False, keys=())

    try:
        lines = path.read_text(encoding="utf-8-sig", errors="replace").splitlines()
    except OSError:
        return SecretFileStatus(exists=True, keys=(), read_error=True)

    found: dict[str, KeyStatus] = {}
    for raw_line in lines:
        line = raw_line.strip()
        if not line or line.startswith(("#", ";", "[")):
            continue
        if line.lower().startswith("export "):
            line = line[7:].lstrip()
        if "=" not in line:
            continue

        key, raw_value = line.split("=", 1)
        key = key.strip().strip("\"'")
        if not is_gemini_key_name(key):
            continue

        value = normalize_secret_value(raw_value)
        found[key] = KeyStatus(
            name=key,
            has_value=bool(value),
            masked_value=mask_secret(value),
        )

    return SecretFileStatus(
        exists=True,
        keys=tuple(sorted(found.values(), key=lambda item: item.name)),
    )


def format_size(size_bytes: int) -> str:
    size_mb = size_bytes / (1024 * 1024)
    if size_mb >= 1024:
        return f"{size_mb / 1024:.2f} GB"
    return f"{size_mb:.2f} MB"


def format_path_list(
    paths: list[Path],
    size_map: dict[Path, int] | None = None,
) -> list[str]:
    if not paths:
        return ["- 없음"]
    result: list[str] = []
    for path in sorted(paths, key=relative_text):
        suffix = ""
        if size_map is not None:
            suffix = f" ({format_size(size_map[path])})"
        result.append(f"- {relative_text(path)}{suffix}")
    return result


def describe_key_status_for_report(status: SecretFileStatus) -> list[str]:
    if not status.exists:
        return ["- 파일 없음"]
    if status.read_error:
        return ["- 파일 존재 / 읽기 실패"]
    if not status.keys:
        return ["- 파일 존재 / Gemini API Key 항목 없음"]
    return [
        f"- {item.name}: {'값 있음' if item.has_value else '비어 있음'}"
        for item in status.keys
    ]


def describe_key_status_for_console(status: SecretFileStatus) -> str:
    if not status.exists:
        return "파일 없음"
    if status.read_error:
        return "파일 존재 / 읽기 실패"
    if not status.keys:
        return "Gemini API Key 항목 없음"
    return ", ".join(
        f"{item.name}={item.masked_value}"
        for item in status.keys
    )


def is_backup_file(path: Path) -> bool:
    name = path.name.lower()
    return (
        name.startswith("app_backup_")
        or name.startswith("app.py.bak")
        or fnmatch.fnmatch(name, "evaluation_template_backup_*.tsv")
        or ("backup" in name and path.suffix.lower() in {".py", ".tsv", ".xlsx"})
    )


def main() -> int:
    files = collect_project_files()
    ignore_patterns = read_gitignore_patterns()
    missing_ignore_patterns = [
        pattern
        for pattern in REQUIRED_IGNORE_PATTERNS
        if pattern not in ignore_patterns
    ]

    env_status = inspect_secret_file(ENV_PATH)
    secrets_status = inspect_secret_file(SECRETS_PATH)
    env_protected = is_ignored(".env", ignore_patterns)
    secrets_protected = is_ignored(
        ".streamlit/secrets.toml",
        ignore_patterns,
    )

    existing_core = [
        PROJECT_ROOT / relative_path
        for relative_path in CORE_FILES
        if (PROJECT_ROOT / relative_path).is_file()
    ]
    missing_core = [
        relative_path
        for relative_path in CORE_FILES
        if not (PROJECT_ROOT / relative_path).is_file()
    ]

    backup_files = [path for path in files if is_backup_file(path)]
    cache_files = [
        path
        for path in files
        if "__pycache__" in path.parts or path.suffix.lower() == ".pyc"
    ]
    temporary_files = [
        path
        for path in files
        if (
            path.name.startswith("~$")
            and path.suffix.lower() == ".xlsx"
        )
        or path.suffix.lower() == ".tmp"
    ]
    unprotected_backup_files = [
        path
        for path in backup_files
        if not is_ignored(relative_text(path), ignore_patterns)
    ]
    unprotected_temporary_files = [
        path
        for path in temporary_files
        if not is_ignored(relative_text(path), ignore_patterns)
    ]

    file_sizes: dict[Path, int] = {}
    for path in files:
        try:
            file_sizes[path] = path.stat().st_size
        except OSError:
            file_sizes[path] = 0

    files_over_50mb = [
        path
        for path, size in file_sizes.items()
        if size >= 50 * 1024 * 1024
    ]
    files_over_100mb = [
        path
        for path, size in file_sizes.items()
        if size >= 100 * 1024 * 1024
    ]

    sensitive_protection_ok = (
        (not env_status.exists or env_protected)
        and (not secrets_status.exists or secrets_protected)
    )
    has_valid_api_key = env_status.has_valid_key or secrets_status.has_valid_key

    share_ready = not (
        missing_ignore_patterns
        or missing_core
        or files_over_100mb
        or not sensitive_protection_ok
        or not has_valid_api_key
        or unprotected_backup_files
        or unprotected_temporary_files
    )
    final_judgment = "공유 준비 가능" if share_ready else "공유 전 확인 필요"

    report_lines = [
        "GitHub 공유 준비 최종 점검 보고서",
        "=" * 72,
        f"점검 시간: {datetime.now().astimezone().strftime('%Y-%m-%d %H:%M:%S %Z')}",
        f"프로젝트 경로: {PROJECT_ROOT}",
        f"전체 파일 수: {len(files)}개",
        f"최종 판단: {final_judgment}",
        "",
        "[핵심 공유 파일]",
        *format_path_list(existing_core),
        "",
        "[누락된 핵심 파일]",
        *([f"- {path}" for path in missing_core] or ["- 없음"]),
        "",
        "[README.md]",
        f"- 존재 여부: {'예' if (PROJECT_ROOT / 'README.md').is_file() else '아니오'}",
        "",
        "[final 비교 Excel]",
        (
            "- 존재 여부: "
            f"{'예' if (PROJECT_ROOT / '12_compare_experiment' / 'llm_rag_comparison_result_final.xlsx').is_file() else '아니오'}"
        ),
        "",
        "[.env 점검]",
        f"- 파일 존재 여부: {'예' if env_status.exists else '아니오'}",
        *describe_key_status_for_report(env_status),
        f"- GitHub 보호 여부: {'예' if env_protected else '아니오'}",
        "",
        "[.streamlit/secrets.toml 점검]",
        f"- 파일 존재 여부: {'예' if secrets_status.exists else '아니오'}",
        *describe_key_status_for_report(secrets_status),
        f"- GitHub 보호 여부: {'예' if secrets_protected else '아니오'}",
        "",
        "[.gitignore 제외 패턴]",
        *([f"- {pattern}" for pattern in ignore_patterns] or ["- 없음"]),
        "",
        "[필수 .gitignore 누락 패턴]",
        *([f"- {pattern}" for pattern in missing_ignore_patterns] or ["- 없음"]),
        "",
        "[50MB 이상 파일]",
        *format_path_list(files_over_50mb, file_sizes),
        "",
        "[100MB 이상 파일]",
        *format_path_list(files_over_100mb, file_sizes),
        "",
        "[백업 파일]",
        *format_path_list(backup_files),
        (
            "- GitHub 미보호 백업 파일: "
            f"{len(unprotected_backup_files)}개"
        ),
        "",
        "[임시 파일]",
        *format_path_list(temporary_files),
        (
            "- GitHub 미보호 임시 파일: "
            f"{len(unprotected_temporary_files)}개"
        ),
        "",
        "[Python 캐시 파일]",
        *format_path_list(cache_files),
        "",
        "[보안 안내]",
        "- 이 보고서에는 API Key 원문 또는 마스킹된 키 값도 저장하지 않습니다.",
        "- 민감 파일은 존재 여부, 키 항목 존재 여부, 값 유무만 기록합니다.",
        "",
        f"최종 판단: {final_judgment}",
    ]

    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text("\n".join(report_lines) + "\n", encoding="utf-8")

    print(
        "README.md 확인 완료: "
        f"{'존재' if (PROJECT_ROOT / 'README.md').is_file() else '누락'}"
    )
    print(
        "final 비교 엑셀 확인 완료: "
        f"{'존재' if (PROJECT_ROOT / '12_compare_experiment' / 'llm_rag_comparison_result_final.xlsx').is_file() else '누락'}"
    )
    print(
        ".env 확인 완료: "
        f"{describe_key_status_for_console(env_status)}"
    )
    print(
        "secrets.toml 확인 완료: "
        f"{describe_key_status_for_console(secrets_status)}"
    )
    print(
        ".gitignore 보호 확인 완료: "
        f".env={'보호' if env_protected else '미보호'}, "
        f"secrets.toml={'보호' if secrets_protected else '미보호'}"
    )
    print(f"50MB 이상 파일: {len(files_over_50mb)}개")
    print(f"100MB 이상 파일: {len(files_over_100mb)}개")
    print(f"최종 판단: {final_judgment}")
    print(f"보고서 저장 경로: {REPORT_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
