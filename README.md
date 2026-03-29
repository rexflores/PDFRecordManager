# PDF Record Manager

Desktop app for organizing, merging, and previewing PDF workflows with a modernized Tkinter UI.

## Release Artifacts

Each official release publishes two Windows distributions:

- Installer: PDFRecordManager-Setup.exe
- Portable: PDFRecordManager-Portable.zip (contains PDFRecordManager-Portable/PDFRecordManager.exe)

## Local Build Commands

- `build_onedir.bat`: build folder-based executable
- `build_onefile.bat`: build single executable
- `build_installer.bat`: build installer (requires Inno Setup)
- `build_portable.bat`: build portable package
- `build_release.bat`: build installer + portable together

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
