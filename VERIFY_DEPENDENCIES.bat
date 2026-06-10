@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=py -3"
) else (
    set "PYTHON_CMD=python"
)

%PYTHON_CMD% -c "import install_dependencies; raise SystemExit(install_dependencies.verify_imports())"
pause
