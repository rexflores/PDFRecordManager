@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

if not defined PDF_AUTOTOOL_PYTHON (
    for %%I in (python.exe) do set "PDF_AUTOTOOL_PYTHON=%%~$PATH:I"
)

if not defined PDF_AUTOTOOL_PYTHON (
    echo Python executable not found in PATH.
    echo Install Python 3.11+ or set PDF_AUTOTOOL_PYTHON to a valid python.exe path.
    exit /b 1
)

if not exist "!PDF_AUTOTOOL_PYTHON!" (
    echo Python executable not found:
    echo   !PDF_AUTOTOOL_PYTHON!
    echo Set PDF_AUTOTOOL_PYTHON to a valid python.exe path.
    exit /b 1
)

set "DEFAULT_ISCC_EXE=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
set "ALT_ISCC_EXE=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
set "ISCC_EXE="

if defined PDF_AUTOTOOL_ISCC (
    set "ISCC_EXE=!PDF_AUTOTOOL_ISCC!"
) else (
    if exist "%DEFAULT_ISCC_EXE%" (
        set "ISCC_EXE=!DEFAULT_ISCC_EXE!"
    ) else if exist "%ALT_ISCC_EXE%" (
        set "ISCC_EXE=!ALT_ISCC_EXE!"
    ) else (
        where iscc >nul 2>&1
        if not errorlevel 1 (
            set "ISCC_EXE=iscc"
        )
    )
)

if not defined ISCC_EXE (
    echo Inno Setup compiler not found.
    echo Install Inno Setup 6 or set PDF_AUTOTOOL_ISCC to your ISCC.exe path.
    exit /b 1
)

if /I not "!ISCC_EXE!"=="iscc" (
    if not exist "!ISCC_EXE!" (
        echo Inno Setup compiler not found:
        echo   !ISCC_EXE!
        echo Set PDF_AUTOTOOL_ISCC to a valid ISCC.exe path.
        exit /b 1
    )
)

echo Using Python: !PDF_AUTOTOOL_PYTHON!
echo Using ISCC: !ISCC_EXE!
echo.
echo [1/2] Building onedir application files...
call "%~dp0build_onedir.bat"
if errorlevel 1 (
    echo.
    echo Installer build stopped: onedir build failed.
    exit /b 1
)

echo.
echo [2/2] Building installer...
if /I "!ISCC_EXE!"=="iscc" (
    iscc "%~dp0installer\PDFRecordManager.iss"
) else (
    "!ISCC_EXE!" "%~dp0installer\PDFRecordManager.iss"
)
if errorlevel 1 (
    echo.
    echo Installer build failed.
    exit /b 1
)

echo.
echo Installer build complete.
echo Output: dist\installer\PDFRecordManager-Setup.exe
exit /b 0
