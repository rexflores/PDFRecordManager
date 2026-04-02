import argparse
import json
import os
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
MAIN_FILE = ROOT_DIR / "main.py"
INSTALLER_FILE = ROOT_DIR / "installer" / "PDFRecordManager.iss"
BUILD_INFO_FILE = ROOT_DIR / "build_info.json"

APP_VERSION_PATTERN = re.compile(
    r'^(?P<prefix>\s*APP_VERSION\s*=\s*")(?P<value>[^"]*)(?P<suffix>".*)$',
    re.MULTILINE,
)
UPDATE_URL_PATTERN = re.compile(
    r'^(?P<prefix>\s*DEFAULT_UPDATE_MANIFEST_URL\s*=\s*")(?P<value>[^"]*)(?P<suffix>".*)$',
    re.MULTILINE,
)
INSTALLER_VERSION_PATTERN = re.compile(
    r'^(?P<prefix>\s*#define\s+MyAppVersion\s+")(?P<value>[^"]*)(?P<suffix>".*)$',
    re.MULTILINE,
)


def _normalize_metadata_value(value):
    text = str(value or "").strip()
    if not text:
        return ""
    if text.lower() in {"unknown", "none", "null", "n/a"}:
        return ""
    return text


def _run_git_text_command(command):
    kwargs = {
        "stdout": subprocess.PIPE,
        "stderr": subprocess.PIPE,
        "text": True,
        "check": False,
        "cwd": ROOT_DIR,
    }
    create_no_window = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    if create_no_window:
        kwargs["creationflags"] = create_no_window

    try:
        result = subprocess.run(command, **kwargs)
    except Exception:
        return ""

    if result.returncode != 0:
        return ""
    return str(result.stdout or "").strip()


def _resolve_build_commit(explicit_commit=None):
    from_args = _normalize_metadata_value(explicit_commit)
    if from_args:
        return from_args

    from_env = _normalize_metadata_value(os.environ.get("PDF_AUTOTOOL_COMMIT"))
    if from_env:
        return from_env

    github_sha = _normalize_metadata_value(os.environ.get("GITHUB_SHA"))
    if github_sha:
        return github_sha[:12]

    from_git = _normalize_metadata_value(
        _run_git_text_command(["git", "rev-parse", "--short=12", "HEAD"])
    )
    if from_git:
        return from_git

    return "unknown"


def _resolve_build_date(explicit_date=None):
    from_args = _normalize_metadata_value(explicit_date)
    if from_args:
        return from_args

    from_env = _normalize_metadata_value(os.environ.get("PDF_AUTOTOOL_BUILD_DATE"))
    if from_env:
        return from_env

    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _write_build_info_file(commit_value, build_date_value):
    payload = {
        "commit": str(commit_value or "unknown"),
        "build_date": str(build_date_value or "unknown"),
    }
    BUILD_INFO_FILE.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    print(f"Stamped build info: commit={payload['commit']}, build_date={payload['build_date']}")


def _replace_single_value(text, pattern, new_value, label, file_path):
    matches = list(pattern.finditer(text))
    if len(matches) != 1:
        raise RuntimeError(
            f"Expected exactly one '{label}' entry in {file_path}, found {len(matches)}."
        )

    match = matches[0]
    old_value = match.group("value")
    if old_value == new_value:
        return text, False

    start, end = match.span("value")
    updated_text = text[:start] + new_value + text[end:]
    print(f"Updated {label}: {old_value} -> {new_value}")
    return updated_text, True


def _validate_version(version):
    if not re.match(r"^\d+\.\d+\.\d+([\-+\.][0-9A-Za-z.-]+)?$", version):
        raise ValueError(
            f"Invalid version '{version}'. Expected format like 1.2.3 or 1.2.3-beta1."
        )


def main():
    parser = argparse.ArgumentParser(
        description="Synchronize release metadata across main.py and installer script."
    )
    parser.add_argument("--version", required=True, help="Release version (example: 1.2.3)")
    parser.add_argument(
        "--update-url",
        default=None,
        help="Default update feed URL to embed in the app build",
    )
    parser.add_argument(
        "--build-commit",
        default=None,
        help="Optional build commit hash to stamp into build_info.json",
    )
    parser.add_argument(
        "--build-date",
        default=None,
        help="Optional build timestamp (ISO-8601) to stamp into build_info.json",
    )
    args = parser.parse_args()

    try:
        _validate_version(args.version)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if not MAIN_FILE.exists() or not INSTALLER_FILE.exists():
        print("Expected project files were not found.", file=sys.stderr)
        return 1

    main_text = MAIN_FILE.read_text(encoding="utf-8")
    installer_text = INSTALLER_FILE.read_text(encoding="utf-8")

    changed = False

    main_text, did_change = _replace_single_value(
        main_text,
        APP_VERSION_PATTERN,
        args.version,
        "APP_VERSION",
        MAIN_FILE,
    )
    changed = changed or did_change

    if args.update_url is not None:
        main_text, did_change = _replace_single_value(
            main_text,
            UPDATE_URL_PATTERN,
            args.update_url,
            "DEFAULT_UPDATE_MANIFEST_URL",
            MAIN_FILE,
        )
        changed = changed or did_change

    installer_text, did_change = _replace_single_value(
        installer_text,
        INSTALLER_VERSION_PATTERN,
        args.version,
        "MyAppVersion",
        INSTALLER_FILE,
    )
    changed = changed or did_change

    if changed:
        MAIN_FILE.write_text(main_text, encoding="utf-8")
        INSTALLER_FILE.write_text(installer_text, encoding="utf-8")
        print("Release metadata synchronized successfully.")
    else:
        print("No metadata changes were required.")

    build_commit = _resolve_build_commit(args.build_commit)
    build_date = _resolve_build_date(args.build_date)
    _write_build_info_file(build_commit, build_date)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
