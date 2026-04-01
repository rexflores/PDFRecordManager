@echo off
setlocal EnableExtensions
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
echo PDF Record Manager - Build Portable Package
echo ===============================================
echo.

echo [1/2] Building onedir application files...
call "%~dp0build_onedir.bat"
if errorlevel 1 (
    echo.
    echo Portable build stopped: onedir build failed.
    exit /b 1
)

set "PORTABLE_DIR=dist\portable"
set "PORTABLE_APP_DIR=%PORTABLE_DIR%\PDFRecordManager-Portable"
set "PORTABLE_EXE=%PORTABLE_APP_DIR%\PDFRecordManager.exe"
set "PORTABLE_README=%PORTABLE_DIR%\README-Portable.txt"
set "PORTABLE_ZIP=%PORTABLE_DIR%\PDFRecordManager-Portable.zip"

if not exist "%PORTABLE_DIR%" mkdir "%PORTABLE_DIR%"
if exist "%PORTABLE_APP_DIR%" rmdir /s /q "%PORTABLE_APP_DIR%"

echo [2/2] Packaging portable files...
robocopy "dist\PDFRecordManager" "%PORTABLE_APP_DIR%" /E /NFL /NDL /NJH /NJS /NC /NS >nul
if errorlevel 8 (
    echo Failed to copy onedir output into portable package folder.
    exit /b 1
)

if not exist "%PORTABLE_EXE%" (
    echo Portable executable was not found at:
    echo   %PORTABLE_EXE%
    exit /b 1
)

(
echo PDF Record Manager Portable
echo ===========================
echo.
echo Run PDFRecordManager-Portable\PDFRecordManager.exe directly.
echo This build does not require installation.
echo.
echo Notes:
echo - Settings are stored per user in %%APPDATA%%\PDF_AutoTool\settings.json
echo - For automatic updates, configure Help ^> Set Update Feed URL
echo - If Windows warns about unknown publisher, that is reputation-based
echo   and can be reduced by code-signing future releases.
) > "%PORTABLE_README%"

if exist "%PORTABLE_ZIP%" del /q "%PORTABLE_ZIP%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path '%PORTABLE_APP_DIR%','%PORTABLE_README%' -DestinationPath '%PORTABLE_ZIP%' -Force"
if errorlevel 1 (
    echo Failed to create portable ZIP package.
    exit /b 1
)

echo.
echo Portable build complete.
echo Folder: dist\portable
echo EXE   : dist\portable\PDFRecordManager-Portable\PDFRecordManager.exe
echo ZIP   : dist\portable\PDFRecordManager-Portable.zip
exit /b 0
