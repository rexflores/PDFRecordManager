import argparse
import re
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
MAIN_FILE = ROOT_DIR / "main.py"
INSTALLER_FILE = ROOT_DIR / "installer" / "PDFRecordManager.iss"

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

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
