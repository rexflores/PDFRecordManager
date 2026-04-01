@echo off
setlocal
cd /d "%~dp0\..\.."

if not defined PDF_AUTOTOOL_PYTHON (
    for %%I in (python.exe) do set "PDF_AUTOTOOL_PYTHON=%%~$PATH:I"
)

if not defined PDF_AUTOTOOL_PYTHON (
    echo Python executable not found in PATH.
    echo Install Python 3.11+ or set PDF_AUTOTOOL_PYTHON to a valid python.exe path.
    exit /b 1
)

if not exist "%PDF_AUTOTOOL_PYTHON%" (
    echo Python executable not found:
    echo   %PDF_AUTOTOOL_PYTHON%
    echo Set PDF_AUTOTOOL_PYTHON to a valid python.exe path.
    exit /b 1
)

echo Using Python: %PDF_AUTOTOOL_PYTHON%
echo.
echo ===============================================
echo PDF Record Manager - Build Release Artifacts
echo ===============================================
echo.

if defined PDF_AUTOTOOL_VERSION (
    echo [0/2] Syncing release metadata...
    if not exist "%~dp0..\set_release_metadata.py" (
        echo Missing metadata sync script: scripts\set_release_metadata.py
        exit /b 1
    )

    if defined PDF_AUTOTOOL_UPDATE_FEED_URL (
        "%PDF_AUTOTOOL_PYTHON%" "%~dp0..\set_release_metadata.py" --version "%PDF_AUTOTOOL_VERSION%" --update-url "%PDF_AUTOTOOL_UPDATE_FEED_URL%"
    ) else (
        "%PDF_AUTOTOOL_PYTHON%" "%~dp0..\set_release_metadata.py" --version "%PDF_AUTOTOOL_VERSION%"
    )

    if errorlevel 1 (
        echo.
        echo Release build stopped: metadata sync failed.
        exit /b 1
    )
    echo.
)

echo [1/2] Building installer package...
call "%~dp0build_installer.bat"
if errorlevel 1 (
    echo.
    echo Release build stopped: installer build failed.
    exit /b 1
)

echo.
echo [2/2] Building portable package...
call "%~dp0build_portable.bat"
if errorlevel 1 (
    echo.
    echo Release build stopped: portable build failed.
    exit /b 1
)

echo.
echo Release artifacts completed successfully.
echo Installer: dist\installer\PDFRecordManager-Setup.exe
echo Portable : dist\portable\PDFRecordManager-Portable.zip
echo Portable EXE: dist\portable\PDFRecordManager-Portable\PDFRecordManager.exe
exit /b 0
