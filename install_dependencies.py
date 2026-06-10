#!/usr/bin/env python3
"""Install and verify Mod Automation dependencies.

This script installs the packages listed in requirements.txt into the Python
interpreter used to run it. On Windows, it also confirms whether pywin32 can be
imported for Excel COM conversion of .xlsb/.xls J.1 files.
"""
from __future__ import annotations

import importlib
import platform
import subprocess
import sys
from pathlib import Path

REQUIRED_IMPORTS = {
    "pandas": "pandas",
    "numpy": "numpy",
    "openpyxl": "openpyxl",
    "pyxlsb": "pyxlsb",
    "xlrd": "xlrd",
    "docx": "python-docx",
}

WINDOWS_ONLY_IMPORTS = {
    "win32com.client": "pywin32",
}


def run_command(command: list[str]) -> None:
    print("\n> " + " ".join(command))
    subprocess.check_call(command)


def import_ok(module_name: str) -> bool:
    try:
        importlib.import_module(module_name)
        return True
    except Exception:
        return False


def verify_imports() -> int:
    print("\nVerifying installed packages...")
    failures: list[str] = []

    for module_name, package_name in REQUIRED_IMPORTS.items():
        if import_ok(module_name):
            print(f"  OK   {package_name}")
        else:
            print(f"  FAIL {package_name} (could not import {module_name})")
            failures.append(package_name)

    if platform.system() == "Windows":
        for module_name, package_name in WINDOWS_ONLY_IMPORTS.items():
            if import_ok(module_name):
                print(f"  OK   {package_name} (Windows Excel conversion support)")
            else:
                print(f"  FAIL {package_name} (needed for .xlsb/.xls to .xlsx conversion)")
                failures.append(package_name)
    else:
        print("  SKIP pywin32 check; it is only used on Windows for Excel conversion")

    if failures:
        print("\nDependency verification failed for:")
        for item in failures:
            print(f"  - {item}")
        return 1

    print("\nAll required dependencies are available.")
    return 0


def main() -> int:
    project_dir = Path(__file__).resolve().parent
    requirements = project_dir / "requirements.txt"

    if not requirements.exists():
        print(f"ERROR: requirements.txt not found at {requirements}")
        return 1

    print("Mod Automation dependency installer")
    print("Python executable:", sys.executable)
    print("Python version:", sys.version.replace("\n", " "))
    print("Project folder:", project_dir)

    try:
        run_command([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
        run_command([sys.executable, "-m", "pip", "install", "-r", str(requirements)])
    except subprocess.CalledProcessError as exc:
        print(f"\nERROR: pip install failed with exit code {exc.returncode}")
        print("If this is a locked-down work machine, run this from the approved Python environment or ask IT to allow these packages.")
        return exc.returncode or 1

    return verify_imports()


if __name__ == "__main__":
    raise SystemExit(main())
