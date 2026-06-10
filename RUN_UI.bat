@echo off
setlocal
cd /d "%~dp0"
python mod_automation_ui.py
if errorlevel 1 (
  echo.
  echo The UI did not start. Confirm Python is installed and available as "python".
  pause
)
