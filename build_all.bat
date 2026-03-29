@echo off
setlocal
cd /d "%~dp0"

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

echo ===============================================
echo PDF Record Manager - Build All (onedir + onefile)
echo ===============================================
echo.

echo [1/2] Building onedir package...
call "%~dp0build_onedir.bat"
if errorlevel 1 (
    echo.
    echo Build-all stopped: onedir build failed.
    exit /b 1
)

echo.
echo [2/2] Building onefile package...
call "%~dp0build_onefile.bat"
if errorlevel 1 (
    echo.
    echo Build-all stopped: onefile build failed.
    exit /b 1
)

echo.
echo All builds completed successfully.
echo Onedir output : dist\PDFRecordManager\PDFRecordManager.exe
echo Onefile output: dist\PDFRecordManager.exe
exit /b 0
