import argparse
import json
import re
import sys
from pathlib import Path


def _validate_version(version):
    if not re.match(r"^\d+\.\d+\.\d+([\-+\.][0-9A-Za-z.-]+)?$", version):
        raise ValueError(
            f"Invalid version '{version}'. Expected format like 1.2.3 or 1.2.3-beta1."
        )


def _validate_repository(repository):
    if not re.match(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$", repository):
        raise ValueError(
            f"Invalid repository '{repository}'. Expected format <owner>/<repo>."
        )


def main():
    parser = argparse.ArgumentParser(
        description="Generate update feed JSON for GitHub Releases."
    )
    parser.add_argument("--version", required=True, help="Release version")
    parser.add_argument(
        "--repository", required=True, help="GitHub repository in <owner>/<repo> format"
    )
    parser.add_argument("--output", required=True, help="Output JSON file path")
    parser.add_argument(
        "--notes",
        default="",
        help="Optional release note text included in the feed",
    )
    args = parser.parse_args()

    try:
        _validate_version(args.version)
        _validate_repository(args.repository)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    release_base = f"https://github.com/{args.repository}/releases/latest"
    feed = {
        "version": args.version,
        "installer_url": f"{release_base}/download/PDFRecordManager-Setup.exe",
        "portable_url": f"{release_base}/download/PDFRecordManager-Portable.zip",
        "release_page_url": release_base,
    }

    if args.notes:
        feed["notes"] = args.notes

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(feed, indent=2) + "\n", encoding="utf-8")

    print(f"Wrote update feed: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
