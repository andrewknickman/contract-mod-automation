"""
Shared file-selection helpers for the mod automation scripts.

These helpers intentionally avoid project-specific hard-coded paths. Each script
can receive paths from command-line arguments, and when an argument is omitted it
falls back to a file/folder picker. If Tkinter is unavailable, the helper falls
back to a console prompt.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, Optional
import sys

EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xlsb", ".xls")
WORD_EXTENSIONS = (".docx",)
EXCEL_FILETYPES = [
    ("Excel files", "*.xlsx *.xlsm *.xlsb *.xls"),
    ("All files", "*.*"),
]
WORD_FILETYPES = [("Word documents", "*.docx"), ("All files", "*.*")]


def _as_path(value: str | Path | None) -> Optional[Path]:
    return Path(value).expanduser().resolve() if value else None


def _root():
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        return root
    except Exception:
        return None


def _console_prompt(prompt: str, must_exist: bool, is_dir: bool = False) -> Path:
    while True:
        value = input(f"{prompt}: ").strip().strip('"')
        path = Path(value).expanduser().resolve()
        if not value:
            print("A path is required.")
            continue
        if must_exist and not path.exists():
            print(f"Path does not exist: {path}")
            continue
        if must_exist and is_dir and not path.is_dir():
            print(f"Expected a folder, got: {path}")
            continue
        if must_exist and not is_dir and not path.is_file():
            print(f"Expected a file, got: {path}")
            continue
        return path


def choose_existing_file(
    value: str | Path | None,
    title: str,
    filetypes=EXCEL_FILETYPES,
) -> Path:
    """Return an existing file from CLI value, file dialog, or console prompt."""
    path = _as_path(value)
    if path:
        if not path.is_file():
            raise FileNotFoundError(f"{title} not found: {path}")
        return path

    root = _root()
    if root:
        from tkinter import filedialog
        try:
            selected = filedialog.askopenfilename(title=title, filetypes=filetypes)
            if selected:
                return Path(selected).resolve()
        finally:
            root.destroy()

    return _console_prompt(title, must_exist=True, is_dir=False)


def choose_optional_file(
    value: str | Path | None,
    title: str,
    filetypes=EXCEL_FILETYPES,
) -> Optional[Path]:
    """Return an optional existing file; blank/cancel means None."""
    path = _as_path(value)
    if path:
        if not path.is_file():
            raise FileNotFoundError(f"{title} not found: {path}")
        return path

    root = _root()
    if root:
        from tkinter import filedialog
        try:
            selected = filedialog.askopenfilename(title=title, filetypes=filetypes)
            return Path(selected).resolve() if selected else None
        finally:
            root.destroy()

    value = input(f"{title} (optional; press Enter to skip): ").strip().strip('"')
    if not value:
        return None
    path = Path(value).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(f"{title} not found: {path}")
    return path


def choose_existing_dir(value: str | Path | None, title: str) -> Path:
    """Return an existing folder from CLI value, folder dialog, or console prompt."""
    path = _as_path(value)
    if path:
        if not path.is_dir():
            raise NotADirectoryError(f"{title} not found: {path}")
        return path

    root = _root()
    if root:
        from tkinter import filedialog
        try:
            selected = filedialog.askdirectory(title=title)
            if selected:
                return Path(selected).resolve()
        finally:
            root.destroy()

    return _console_prompt(title, must_exist=True, is_dir=True)


def choose_output_dir(value: str | Path | None, title: str) -> Path:
    """Return a folder and create it when necessary."""
    path = _as_path(value)
    if path:
        path.mkdir(parents=True, exist_ok=True)
        return path

    root = _root()
    if root:
        from tkinter import filedialog
        try:
            selected = filedialog.askdirectory(title=title, mustexist=False)
            if selected:
                path = Path(selected).resolve()
                path.mkdir(parents=True, exist_ok=True)
                return path
        finally:
            root.destroy()

    path = _console_prompt(title, must_exist=False, is_dir=True)
    path.mkdir(parents=True, exist_ok=True)
    return path


def choose_save_file(
    value: str | Path | None,
    title: str,
    default_name: str,
    filetypes=EXCEL_FILETYPES,
) -> Path:
    """Return a save path from CLI value, save dialog, or console prompt."""
    path = _as_path(value)
    if path:
        path.parent.mkdir(parents=True, exist_ok=True)
        return path

    root = _root()
    if root:
        from tkinter import filedialog
        try:
            selected = filedialog.asksaveasfilename(
                title=title,
                initialfile=default_name,
                filetypes=filetypes,
            )
            if selected:
                path = Path(selected).resolve()
                path.parent.mkdir(parents=True, exist_ok=True)
                return path
        finally:
            root.destroy()

    while True:
        value = input(f"{title} [{default_name}]: ").strip().strip('"')
        if not value:
            value = default_name
        path = Path(value).expanduser().resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        return path


def iter_excel_files(folder: Path, exclude_substrings: Iterable[str] = ()): 
    """Yield non-temp Excel files in a folder, supporting xlsx/xlsm/xlsb/xls."""
    excludes = tuple(s.lower() for s in exclude_substrings)
    for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        lower = path.name.lower()
        if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS:
            if lower.startswith("~$"):
                continue
            if any(ex in lower for ex in excludes):
                continue
            yield path


def find_excel_by_prefix(folder: Path, prefix: str) -> Optional[Path]:
    """Find the first Excel file whose name starts with prefix, case-insensitive."""
    wanted = prefix.lower()
    for path in iter_excel_files(folder):
        if path.name.lower().startswith(wanted):
            return path
    for path in iter_excel_files(folder):
        if wanted in path.name.lower():
            return path
    return None


def read_excel_auto(path: str | Path, *args, **kwargs):
    """pd.read_excel wrapper that enables pyxlsb for .xlsb files."""
    import pandas as pd
    path = Path(path)
    if path.suffix.lower() == ".xlsb" and "engine" not in kwargs:
        kwargs["engine"] = "pyxlsb"
    return pd.read_excel(path, *args, **kwargs)


def excel_file(path: str | Path):
    """pd.ExcelFile wrapper that enables pyxlsb for .xlsb files."""
    import pandas as pd
    path = Path(path)
    if path.suffix.lower() == ".xlsb":
        return pd.ExcelFile(path, engine="pyxlsb")
    return pd.ExcelFile(path)
