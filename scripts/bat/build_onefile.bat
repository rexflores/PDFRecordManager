@echo off
setlocal
cd /d "%~dp0\..\.."

set "PYTHON_EXE="
if defined PDF_AUTOTOOL_PYTHON (
    set "PYTHON_EXE=%PDF_AUTOTOOL_PYTHON%"
) else (
    for %%I in (python.exe) do set "PYTHON_EXE=%%~$PATH:I"
)

if /I "%PYTHON_EXE%"=="python" (
    for %%I in (python.exe) do set "PYTHON_EXE=%%~$PATH:I"
)

if not defined PYTHON_EXE (
    echo Python executable not found in PATH.
    echo Install Python 3.11+ or set PDF_AUTOTOOL_PYTHON to a valid python.exe path.
    exit /b 1
)

if not exist "%PYTHON_EXE%" (
    echo Python executable not found:
    echo   %PYTHON_EXE%
    echo.
    echo Set PDF_AUTOTOOL_PYTHON to a valid python.exe path.
    exit /b 1
)

echo Using Python: %PYTHON_EXE%
echo Building onefile EXE using PDFRecordManager.onefile.spec...
"%PYTHON_EXE%" -m PyInstaller --clean -y PDFRecordManager.onefile.spec
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

echo Build complete.
echo Output: dist\PDFRecordManager.exe
exit /b 0
