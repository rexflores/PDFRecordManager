# PDF Record Manager

Desktop app for organizing, merging, and previewing PDF workflows with a modernized Tkinter UI.

## Release Artifacts

Each official release publishes two Windows distributions:

- Installer: PDFRecordManager-Setup.exe
- Portable: PDFRecordManager-Portable.zip (contains PDFRecordManager-Portable/PDFRecordManager.exe)

## Local Build Commands

Preferred (no BAT required):

- `python scripts/build.py --target release`: build installer + portable together
- `python scripts/build.py --target installer`: build installer only (requires Inno Setup)
- `python scripts/build.py --target portable`: build portable package only
- `python scripts/build.py --target onedir`: build folder-based executable only
- `python scripts/build.py --target onefile`: build single executable only
- `python scripts/build.py --target all`: build onedir + onefile together

Optional Windows BAT wrappers (kept for convenience):

- `scripts/bat/build_onedir.bat`
- `scripts/bat/build_onefile.bat`
- `scripts/bat/build_installer.bat`
- `scripts/bat/build_portable.bat`
- `scripts/bat/build_release.bat`

## Official GitHub Publishing

This repository includes GitHub automation in `.github/workflows/release.yml`.

What it does:

- Builds installer and portable artifacts on Windows runner
- Publishes assets to GitHub Releases
- Generates and deploys `update-feed.json` to GitHub Pages
- Publishes `update-feed.json` as a GitHub Release asset backup
- Publishes SHA256 checksums for release verification
- Stamps APP_VERSION and installer version from the release version

## Security Notes

- Unsigned Windows executables can trigger SmartScreen or antivirus reputation warnings.
- This risk is lower with onedir portable builds than onefile bundles, but warnings may still appear.
- For official trusted distribution, sign executables and installer with a code-signing certificate.

Trigger options:

- Push a tag: v1.0.1
- Run workflow manually with version input

## Update Feed URL

After first workflow run with Pages enabled, your feed URL is:

`https://raw.githubusercontent.com/OWNER/REPO/gh-pages/update-feed.json`

The release workflow also embeds this URL into the published app build.

Optional: if GitHub Pages is configured, `https://OWNER.github.io/REPO/update-feed.json`
can also serve the same file.

Release asset fallback URL:

`https://github.com/OWNER/REPO/releases/latest/download/update-feed.json`

## First-Time Setup

Follow `PUBLISH_GITHUB.md` for one-time repository setup and first release steps.
