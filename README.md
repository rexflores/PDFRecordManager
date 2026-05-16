# PDF Record Manager

A Windows desktop app for organizing, merging, and reviewing PDF employee records.
It provides fast search, guided save/merge workflows, and batch processing tools with a modernized Tkinter UI.

## Features

- Save new records with structured folder paths and standardized filenames
- Merge pending PDFs into existing records with preview and backup options
- Flexible name and folder search (ignores punctuation and order)
- Batch processing utilities and preview tools
- Optional recycle bin cleanup (uses send2trash when available)
- Persistent preferences and safe confirmation controls

## Quick Start

### Option 1: Download and run (recommended)

1. Go to the GitHub Releases page for this repository.
2. Download one of the artifacts:
   - Installer: PDFRecordManager-Setup.exe
   - Portable: PDFRecordManager-Portable.zip
3. Run the installer or extract the portable zip and launch the EXE.

### Option 2: Run from source

Requirements:
- Windows 10/11
- Python 3.11+

Steps:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
```

## Usage

### Save new record

1. Select the Records Root Folder and Pending Folder.
2. Choose or type the employee name.
3. Fill out year details and status.
4. Save the record; the file is copied to the target folder and the pending file is archived.

### Merge into existing record

1. Select a pending PDF.
2. Choose the destination employee folder.
3. Pick the existing PDF to merge into.
4. Review the summary, then merge and save.

## Settings

Preferences are saved in a local settings file and persist across sessions. You can customize:

- Confirmation prompts for save and merge actions
- Backup behavior when replacing PDFs
- Display preferences (icons and text)

## Build Windows Distributions

These commands generate the installer and portable packages:

```powershell
python scripts/build.py --target release
```

Other build options:

```powershell
python scripts/build.py --target installer
python scripts/build.py --target portable
python scripts/build.py --target onedir
python scripts/build.py --target onefile
python scripts/build.py --target all
```

## Update Workflow

Releases are published through GitHub Actions and a hosted update feed:

```powershell
git tag v1.3.3
git push origin v1.3.3
```

This produces:

- Installer and portable artifacts
- Update feed JSON for in-app checks

See UPDATE_WORKFLOW.md and PUBLISH_GITHUB.md for details.

## Troubleshooting

- If recycle bin actions are unavailable, install send2trash.
- If encryption features fail, ensure cryptography is installed.
- If PDFs fail to open, confirm pypdf (or PyPDF2) is installed.

## Contributing

Pull requests are welcome. Please keep changes focused and include a short description of the behavior change.

## License

MIT License. See [LICENSE](LICENSE).