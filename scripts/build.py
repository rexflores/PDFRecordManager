import argparse
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = ROOT_DIR / "scripts"
SET_METADATA_SCRIPT = SCRIPTS_DIR / "set_release_metadata.py"
ONEDIR_SPEC = ROOT_DIR / "PDFRecordManager.spec"
ONEFILE_SPEC = ROOT_DIR / "PDFRecordManager.onefile.spec"
INSTALLER_FILE = ROOT_DIR / "installer" / "PDFRecordManager.iss"
NUITKA_OUTPUT_DIR = ROOT_DIR / "dist" / "nuitka"
STAGED_RELEASE_DIR = ROOT_DIR / "dist" / "PDFRecordManager"
NIGHTLY_ARTIFACT_DIRS = (
    ROOT_DIR / "build",
    ROOT_DIR / "dist",
)

VERSION_PATTERN = re.compile(r"^\d+\.\d+\.\d+([\-+\.][0-9A-Za-z.-]+)?$")

PORTABLE_README_TEXT = """PDF Record Manager Portable
===========================

Run PDFRecordManager-Portable\\PDFRecordManager.exe directly.
This build does not require installation.

Notes:
- Settings are stored per user in %APPDATA%\\PDF_AutoTool\\settings.json
- For automatic updates, configure Help > Set Update Feed URL
- If Windows warns about unknown publisher, that is reputation-based
  and can be reduced by code-signing future releases.
"""


def _run_command(command, step_label):
    print(step_label)
    print(f"> {subprocess.list2cmdline(command)}")
    result = subprocess.run(command, cwd=ROOT_DIR)
    if result.returncode != 0:
        raise RuntimeError(f"{step_label} failed with exit code {result.returncode}.")


def _run_nuitka_command(python_exe, args, step_label):
    command = [python_exe, "-m", "nuitka"] + list(args)
    _run_command(command, step_label)


def _clean_release_environment():
    for path in NIGHTLY_ARTIFACT_DIRS:
        if path.exists():
            shutil.rmtree(path)


def _resolve_python_executable(python_override=None):
    candidates = []
    if python_override:
        candidates.append(python_override)

    env_python = os.environ.get("PDF_AUTOTOOL_PYTHON")
    if env_python:
        candidates.append(env_python)

    if sys.executable:
        candidates.append(sys.executable)

    candidates.append("python")

    for candidate in candidates:
        candidate_path = Path(candidate)
        if candidate_path.exists():
            return str(candidate_path)

        resolved = shutil.which(candidate)
        if resolved:
            return resolved

    raise RuntimeError(
        "Python executable not found. Install Python 3.11+ or set PDF_AUTOTOOL_PYTHON."
    )


def _resolve_iscc_executable(iscc_override=None):
    candidates = []
    if iscc_override:
        candidates.append(iscc_override)

    env_iscc = os.environ.get("PDF_AUTOTOOL_ISCC")
    if env_iscc:
        candidates.append(env_iscc)

    candidates.extend(
        [
            r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
            os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Inno Setup 6\ISCC.exe"),
            "iscc",
        ]
    )

    for candidate in candidates:
        if not candidate:
            continue

        if candidate.lower() == "iscc":
            resolved = shutil.which("iscc")
            if resolved:
                return resolved
            continue

        candidate_path = Path(candidate)
        if candidate_path.exists():
            return str(candidate_path)

        resolved = shutil.which(candidate)
        if resolved:
            return resolved

    raise RuntimeError(
        "Inno Setup compiler not found. Install Inno Setup 6 or set PDF_AUTOTOOL_ISCC."
    )


def _sync_release_metadata(python_exe, version, update_url=None):
    if not SET_METADATA_SCRIPT.exists():
        raise RuntimeError(f"Missing metadata sync script: {SET_METADATA_SCRIPT}")

    command = [python_exe, str(SET_METADATA_SCRIPT), "--version", version]
    if update_url is not None:
        command.extend(["--update-url", update_url])

    _run_command(command, "[0] Syncing release metadata...")


def _build_onedir(python_exe):
    if not ONEDIR_SPEC.exists():
        raise RuntimeError(f"Missing spec file: {ONEDIR_SPEC}")

    _run_command(
        [python_exe, "-m", "PyInstaller", "--clean", "-y", str(ONEDIR_SPEC)],
        "[1] Building onedir executable...",
    )


def _build_onefile(python_exe, step_label="[1] Building onefile executable..."):
    if not ONEFILE_SPEC.exists():
        raise RuntimeError(f"Missing spec file: {ONEFILE_SPEC}")

    _run_command(
        [python_exe, "-m", "PyInstaller", "--clean", "-y", str(ONEFILE_SPEC)],
        step_label,
    )


def _resolve_nuitka_artifact():
    if not NUITKA_OUTPUT_DIR.exists():
        raise RuntimeError(f"Missing Nuitka output directory: {NUITKA_OUTPUT_DIR}")

    dist_candidates = sorted(
        path for path in NUITKA_OUTPUT_DIR.glob("*.dist") if path.is_dir()
    )
    if not dist_candidates:
        raise RuntimeError(
            f"Unable to locate Nuitka standalone output under {NUITKA_OUTPUT_DIR}."
        )

    return dist_candidates[0]


def _stage_nuitka_release(artifact_path):
    if STAGED_RELEASE_DIR.exists():
        shutil.rmtree(STAGED_RELEASE_DIR)
    STAGED_RELEASE_DIR.mkdir(parents=True, exist_ok=True)

    exe_candidate = None
    for item in artifact_path.iterdir():
        if item.is_file() and item.suffix.lower() == ".exe":
            if item.name.lower() == "pdfrecordmanager.exe":
                exe_candidate = item
                break
            if exe_candidate is None:
                exe_candidate = item

    if exe_candidate is None:
        raise RuntimeError(
            f"Unable to locate the Nuitka executable inside {artifact_path}."
        )

    target_exe = STAGED_RELEASE_DIR / "PDFRecordManager.exe"
    shutil.copy2(exe_candidate, target_exe)

    for item in artifact_path.iterdir():
        if item == exe_candidate:
            continue
        destination = STAGED_RELEASE_DIR / item.name
        if item.is_dir():
            shutil.copytree(item, destination)
        else:
            shutil.copy2(item, destination)


def _build_nuitka_release(python_exe, version=None):
    if version:
        file_version = version
    else:
        file_version = None

    _clean_release_environment()
    NUITKA_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    command_args = [
        "--standalone",
        "--windows-console-mode=disable",
        "--enable-plugin=tk-inter",
        f"--output-dir={NUITKA_OUTPUT_DIR}",
        "--windows-icon-from-ico=application.ico",
        "--include-data-file=application.ico=application.ico",
        "--include-data-file=build_info.json=build_info.json",
        "--include-package=pypdf",
        "--include-package=pdfplumber",
        "--include-package=openpyxl",
    ]

    if file_version:
        command_args.extend(
            [
                f"--file-version={file_version}",
                f"--product-version={file_version}",
                "--product-name=PDF Record Manager",
                "--company-name=PDF Record Manager",
            ]
        )

    command_args.append(str(ROOT_DIR / "main.py"))

    _run_nuitka_command(
        python_exe,
        command_args,
        "[1] Building Nuitka standalone executable...",
    )

    artifact_path = _resolve_nuitka_artifact()
    _stage_nuitka_release(artifact_path)


def _build_installer(iscc_exe):
    if not INSTALLER_FILE.exists():
        raise RuntimeError(f"Missing installer script: {INSTALLER_FILE}")

    _run_command([iscc_exe, str(INSTALLER_FILE)], "[2] Building installer package...")


def _write_portable_archive():
    source_dir = ROOT_DIR / "dist" / "PDFRecordManager"
    if not source_dir.exists():
        raise RuntimeError(f"Expected onedir output not found: {source_dir}")

    portable_dir = ROOT_DIR / "dist" / "portable"
    portable_app_dir = portable_dir / "PDFRecordManager-Portable"
    portable_exe = portable_app_dir / "PDFRecordManager.exe"
    portable_readme = portable_dir / "README-Portable.txt"
    portable_zip = portable_dir / "PDFRecordManager-Portable.zip"

    portable_dir.mkdir(parents=True, exist_ok=True)

    if portable_app_dir.exists():
        shutil.rmtree(portable_app_dir)

    shutil.copytree(source_dir, portable_app_dir)

    if not portable_exe.exists():
        raise RuntimeError(f"Portable executable was not found: {portable_exe}")

    portable_readme.write_text(PORTABLE_README_TEXT, encoding="utf-8")

    if portable_zip.exists():
        portable_zip.unlink()

    with zipfile.ZipFile(portable_zip, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file_path in sorted(portable_app_dir.rglob("*")):
            if file_path.is_file():
                archive.write(
                    file_path,
                    arcname=file_path.relative_to(portable_dir).as_posix(),
                )

        archive.write(portable_readme, arcname=portable_readme.name)


def _print_outputs(target):
    if target == "onedir":
        print("Output: dist\\PDFRecordManager\\PDFRecordManager.exe")
    elif target == "onefile":
        print("Output: dist\\PDFRecordManager.exe")
    elif target == "all":
        print("Onedir output : dist\\PDFRecordManager\\PDFRecordManager.exe")
        print("Onefile output: dist\\PDFRecordManager.exe")
    elif target == "installer":
        print("Output: dist\\installer\\PDFRecordManager-Setup.exe")
    elif target == "portable":
        print("Folder: dist\\portable")
        print("EXE   : dist\\portable\\PDFRecordManager-Portable\\PDFRecordManager.exe")
        print("ZIP   : dist\\portable\\PDFRecordManager-Portable.zip")
    elif target == "nuitka":
        print("Output: dist\\PDFRecordManager\\PDFRecordManager.exe")
    else:
        print("Installer: dist\\installer\\PDFRecordManager-Setup.exe")
        print("Portable : dist\\portable\\PDFRecordManager-Portable.zip")
        print("Portable EXE: dist\\portable\\PDFRecordManager-Portable\\PDFRecordManager.exe")


def main():
    parser = argparse.ArgumentParser(description="Python build entrypoint for PDF Record Manager.")
    parser.add_argument(
        "--target",
        choices=("release", "installer", "portable", "onedir", "onefile", "all"),
        default="release",
        help="Build target to run (default: release).",
    )
    parser.add_argument(
        "--version",
        default=os.environ.get("PDF_AUTOTOOL_VERSION"),
        help="Optional release version used for metadata sync (example: 1.2.3).",
    )
    parser.add_argument(
        "--update-url",
        default=os.environ.get("PDF_AUTOTOOL_UPDATE_FEED_URL"),
        help="Optional update feed URL used during metadata sync.",
    )
    parser.add_argument(
        "--python-exe",
        default=None,
        help="Optional override for python executable path.",
    )
    parser.add_argument(
        "--iscc-exe",
        default=None,
        help="Optional override for ISCC.exe path.",
    )
    parser.add_argument(
        "--backend",
        choices=("pyinstaller", "nuitka"),
        default="pyinstaller",
        help="Build backend to use (default: pyinstaller).",
    )
    args = parser.parse_args()

    try:
        if args.update_url and not args.version:
            raise RuntimeError("--update-url requires --version.")

        if args.version and not VERSION_PATTERN.match(args.version):
            raise RuntimeError(
                f"Invalid version '{args.version}'. Expected format like 1.2.3 or 1.2.3-beta1."
            )

        python_exe = _resolve_python_executable(args.python_exe)
        print(f"Using Python: {python_exe}")

        if args.version:
            _sync_release_metadata(python_exe, args.version, args.update_url)

        if args.backend == "nuitka":
            if args.target == "onefile":
                raise RuntimeError("Nuitka backend does not support onefile builds. Use --target onedir or --target release.")
            _build_nuitka_release(python_exe, args.version)
        elif args.target == "onefile":
            _build_onefile(python_exe)
        elif args.target == "all":
            _build_onedir(python_exe)
            _build_onefile(python_exe, step_label="[2] Building onefile executable...")
        else:
            if args.target in {"release", "installer", "portable", "onedir"}:
                _build_onedir(python_exe)

            if args.target in {"release", "installer"}:
                iscc_exe = _resolve_iscc_executable(args.iscc_exe)
                print(f"Using ISCC: {iscc_exe}")
                _build_installer(iscc_exe)

            if args.target in {"release", "portable"}:
                print("[3] Packaging portable bundle...")
                _write_portable_archive()

        if args.backend == "nuitka":
            if args.target in {"release", "installer", "all"}:
                iscc_exe = _resolve_iscc_executable(args.iscc_exe)
                print(f"Using ISCC: {iscc_exe}")
                _build_installer(iscc_exe)

            if args.target in {"release", "portable", "all"}:
                print("[3] Packaging portable bundle...")
                _write_portable_archive()

        print("Build completed successfully.")
        _print_outputs(args.target)
        return 0
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
