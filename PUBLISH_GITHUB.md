# Publish To GitHub (Official Release Setup)

This guide makes your app officially published through GitHub Releases + GitHub Pages.

## 1) Create GitHub repository and push this project

Run in project root:

```powershell
git init
git add .
git commit -m "Initial publish setup"
git branch -M main
git remote add origin https://github.com/OWNER/REPO.git
git push -u origin main
```

## 2) Enable required repository settings

In GitHub repository settings:

1. Actions > General:
Workflow permissions: Read and write permissions
2. Pages:
Source: Deploy from a branch
Branch: gh-pages
Folder: /(root)

## 3) Create your first official release

Option A (recommended): tag push

```powershell
git tag v1.0.1
git push origin v1.0.1
```

Option B: manual workflow run

- Open Actions > Publish Official Release
- Run workflow and input version (example: 1.0.1)

## 4) Verify release outputs

After workflow success:

- GitHub Release contains:
  - PDFRecordManager-Setup.exe
  - PDFRecordManager-Portable.zip
  - SHA256SUMS.txt
  - update-feed.json
- GitHub Pages contains:
  - update-feed.json

Feed URL format:

`https://raw.githubusercontent.com/OWNER/REPO/gh-pages/update-feed.json`

Optional GitHub Pages alias (if configured):

`https://OWNER.github.io/REPO/update-feed.json`

Release asset fallback URL:

`https://github.com/OWNER/REPO/releases/latest/download/update-feed.json`

## 5) Verify in-app update flow

On test machine:

1. Open app
2. Go to Help > Check for Updates
3. If newer version exists, choose Installer or Portable

## 6) Next releases

For each new release:

1. Push a new tag (example: v1.0.2)
2. Workflow rebuilds artifacts and refreshes update-feed.json automatically

## Notes

- Version is synchronized into both main.py (APP_VERSION) and installer/PDFRecordManager.iss.
- If a release is rerun manually for an existing version, it updates artifacts on the same release tag.
- Unsigned binaries can trigger SmartScreen/AV reputation warnings. For official distribution,
  sign the installer and executable with a trusted code-signing certificate.
