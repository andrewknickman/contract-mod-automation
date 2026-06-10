#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

echo "============================================================"
echo "Mod Automation dependency installer"
echo "============================================================"
echo

if command -v python3 >/dev/null 2>&1; then
  PYTHON_CMD="python3"
else
  PYTHON_CMD="python"
fi

echo "Using Python command: $PYTHON_CMD"
echo
"$PYTHON_CMD" install_dependencies.py

echo
echo "Dependency installation completed successfully."
