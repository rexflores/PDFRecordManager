# PDF Record Manager Update Workflow

This app now supports in-app update checks via a hosted JSON feed.

The official publishing path is GitHub Releases + GitHub Pages, automated by:

- `.github/workflows/release.yml`

## 1) First-time repository setup

1. Push this project to GitHub.
2. In repository settings:
Actions > General > Workflow permissions = Read and write permissions
Pages > Source = Deploy from a branch
Pages > Branch = gh-pages
Pages > Folder = /(root)

Detailed step-by-step commands are in PUBLISH_GITHUB.md.

## 2) Publish an official release

Recommended trigger (tag-based):

```powershell
git tag v1.0.1
git push origin v1.0.1
```

Alternative trigger:

- Run workflow Publish Official Release manually in GitHub Actions.
- Pass version input (example: 1.0.1).

Workflow output:

- GitHub Release assets:
  - dist/installer/PDFRecordManager-Setup.exe
  - dist/portable/PDFRecordManager-Portable.zip
  - release-manifest/SHA256SUMS.txt
  - release-manifest/update-feed.json
- GitHub Pages file:
  - update-feed.json

## 3) Update feed JSON (automatic)

The workflow generates and publishes update-feed.json with this shape:

```json
{
  "version": "1.0.1",
  "installer_url": "https://github.com/OWNER/REPO/releases/latest/download/PDFRecordManager-Setup.exe",
  "portable_url": "https://github.com/OWNER/REPO/releases/latest/download/PDFRecordManager-Portable.zip",
  "release_page_url": "https://github.com/OWNER/REPO/releases/latest",
  "notes": "Release 1.0.1"
}
```

App behavior:

- If installer_url and portable_url exist: app asks user which one to open
- If only one URL exists: app opens that one after confirmation
- If neither exists but release_page_url exists: app offers to open releases page

## 4) Default feed URL

Feed URL format:

`https://raw.githubusercontent.com/OWNER/REPO/gh-pages/update-feed.json`

Optional GitHub Pages alias (if configured):

`https://OWNER.github.io/REPO/update-feed.json`

Release asset fallback URL:

`https://github.com/OWNER/REPO/releases/latest/download/update-feed.json`

The release workflow also syncs this URL into `main.py` as `DEFAULT_UPDATE_MANIFEST_URL` during the CI build.

## 5) User update flow

- User clicks Help > Check for Updates (or auto check runs on startup)
- App compares versions
- If newer version exists, app prompts for installer or portable download
- Installer path: user runs installer to update in place
- Portable path: user downloads zip and runs PDFRecordManager-Portable/PDFRecordManager.exe

## 6) Manual fallback (without GitHub Actions)

Run:

```powershell
build_release.bat
```

Output:

- dist/installer/PDFRecordManager-Setup.exe
- dist/portable/PDFRecordManager-Portable.zip
- dist/portable/PDFRecordManager-Portable/PDFRecordManager.exe

If you only need one target locally:

- Installer only: build_installer.bat
- Portable only: build_portable.bat

## 7) Configure client app manually (only if needed)

In the app:

- Open Help > Set Update Feed URL
- Paste the feed JSON URL

## Notes

- Keep `AppId` unchanged in `installer/PDFRecordManager.iss` so upgrades replace older installs.
- Keep old release artifacts available so existing feed URLs do not break.
- If users report antivirus warnings, prefer installer or onedir portable ZIP and verify `SHA256SUMS.txt`.
- For production trust and lower false positives, sign release binaries with a code-signing certificate.
