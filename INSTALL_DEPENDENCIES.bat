@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Mod Automation dependency installer
echo ============================================================
echo.

where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=py -3"
) else (
    set "PYTHON_CMD=python"
)

echo Using Python command: %PYTHON_CMD%
echo.
%PYTHON_CMD% install_dependencies.py
set "INSTALL_EXIT=%ERRORLEVEL%"

echo.
if %INSTALL_EXIT% EQU 0 (
    echo Dependency installation completed successfully.
) else (
    echo Dependency installation failed with exit code %INSTALL_EXIT%.
)
echo.
pause
exit /b %INSTALL_EXIT%
