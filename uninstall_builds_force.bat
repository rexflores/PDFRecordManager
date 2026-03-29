@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo ===============================================
echo PDF Record Manager - Force Remove Build Artifacts
echo ===============================================
echo Running in force mode (no confirmation prompt).
echo Removing generated files only:
echo   - build\
echo   - dist\
echo   - all __pycache__ folders
echo   - all *.pyc and *.pyo files
echo.

set "REMOVED_ANY=0"

if exist "build" (
    echo Removing build\
    rmdir /S /Q "build"
    set "REMOVED_ANY=1"
)

if exist "dist" (
    echo Removing dist\
    rmdir /S /Q "dist"
    set "REMOVED_ANY=1"
)

for /D /R %%D in (__pycache__) do (
    if exist "%%~fD" (
        echo Removing %%~fD
        rmdir /S /Q "%%~fD"
        set "REMOVED_ANY=1"
    )
)

for /R %%F in (*.pyc) do (
    if exist "%%~fF" (
        echo Deleting %%~fF
        del /F /Q "%%~fF"
        set "REMOVED_ANY=1"
    )
)

for /R %%F in (*.pyo) do (
    if exist "%%~fF" (
        echo Deleting %%~fF
        del /F /Q "%%~fF"
        set "REMOVED_ANY=1"
    )
)

if "!REMOVED_ANY!"=="0" (
    echo No build artifacts were found.
) else (
    echo Force cleanup complete.
)

exit /b 0
