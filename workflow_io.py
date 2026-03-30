from __future__ import annotations

import json
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

STATE_FILE = Path(__file__).with_name('.workflow_last_paths.json')


def _load_state() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding='utf-8'))
        except Exception:
            return {}
    return {}


def _save_state(state: dict) -> None:
    try:
        STATE_FILE.write_text(json.dumps(state, indent=2), encoding='utf-8')
    except Exception:
        pass


def _remember(key: str, value: str) -> None:
    state = _load_state()
    state[key] = value
    _save_state(state)


def _recall(key: str, default: Optional[str] = None) -> Optional[str]:
    state = _load_state()
    return state.get(key, default)


def _normalize_filetypes(filetypes: Optional[Sequence[Tuple[str, str]]]) -> List[Tuple[str, str]]:
    if not filetypes:
        return [('All Files', '*.*')]
    return list(filetypes)


def _pick_with_tk(action: str, *, title: str, filetypes=None, initialdir=None, initialfile=None, mustexist=True):
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    options = {'title': title}
    if initialdir:
        options['initialdir'] = str(initialdir)
    if initialfile:
        options['initialfile'] = str(initialfile)
    if filetypes:
        options['filetypes'] = _normalize_filetypes(filetypes)

    try:
        if action == 'open_file':
            result = filedialog.askopenfilename(**options)
        elif action == 'open_files':
            result = filedialog.askopenfilenames(**options)
        elif action == 'open_dir':
            result = filedialog.askdirectory(**options, mustexist=mustexist)
        elif action == 'save_file':
            result = filedialog.asksaveasfilename(**options)
        else:
            raise ValueError(f'Unsupported action: {action}')
    finally:
        root.destroy()
    return result


def _fallback_input(prompt: str, expect_multiple: bool = False) -> str | List[str]:
    print(prompt)
    if expect_multiple:
        print('Enter one path per line. Submit an empty line when finished.')
        values = []
        while True:
            value = input('> ').strip().strip('"')
            if not value:
                break
            values.append(value)
        return values
    return input('> ').strip().strip('"')


def choose_file(*, title: str, filetypes=None, state_key: Optional[str] = None,
                initialdir: Optional[str | Path] = None) -> Path:
    initialdir = Path(_recall(state_key, str(initialdir) if initialdir else '.')) if state_key else Path(initialdir or '.')
    try:
        result = _pick_with_tk('open_file', title=title, filetypes=filetypes, initialdir=initialdir)
    except Exception:
        result = _fallback_input(f'{title}\nEnter the full file path:')
    if not result:
        raise SystemExit(f'Cancelled: {title}')
    path = Path(result)
    if state_key:
        _remember(state_key, str(path.parent))
    return path


def choose_files(*, title: str, filetypes=None, state_key: Optional[str] = None,
                 initialdir: Optional[str | Path] = None) -> List[Path]:
    initialdir = Path(_recall(state_key, str(initialdir) if initialdir else '.')) if state_key else Path(initialdir or '.')
    try:
        result = _pick_with_tk('open_files', title=title, filetypes=filetypes, initialdir=initialdir)
    except Exception:
        result = _fallback_input(f'{title}\nEnter the full file paths:', expect_multiple=True)
    if not result:
        raise SystemExit(f'Cancelled: {title}')
    paths = [Path(p) for p in result]
    if state_key and paths:
        _remember(state_key, str(paths[0].parent))
    return paths


def choose_directory(*, title: str, state_key: Optional[str] = None,
                     initialdir: Optional[str | Path] = None, mustexist: bool = True) -> Path:
    initialdir = Path(_recall(state_key, str(initialdir) if initialdir else '.')) if state_key else Path(initialdir or '.')
    try:
        result = _pick_with_tk('open_dir', title=title, initialdir=initialdir, mustexist=mustexist)
    except Exception:
        result = _fallback_input(f'{title}\nEnter the full directory path:')
    if not result:
        raise SystemExit(f'Cancelled: {title}')
    path = Path(result)
    if state_key:
        _remember(state_key, str(path))
    return path


def choose_save_file(*, title: str, default_name: str, filetypes=None,
                     state_key: Optional[str] = None,
                     initialdir: Optional[str | Path] = None) -> Path:
    initialdir = Path(_recall(state_key, str(initialdir) if initialdir else '.')) if state_key else Path(initialdir or '.')
    try:
        result = _pick_with_tk('save_file', title=title, filetypes=filetypes, initialdir=initialdir, initialfile=default_name)
    except Exception:
        result = _fallback_input(f'{title}\nEnter the full output file path:')
    if not result:
        raise SystemExit(f'Cancelled: {title}')
    path = Path(result)
    if state_key:
        _remember(state_key, str(path.parent))
    return path


def ask_integer(*, title: str, prompt: str, default: Optional[int] = None) -> int:
    try:
        import tkinter as tk
        from tkinter import simpledialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        value = simpledialog.askinteger(title=title, prompt=prompt, initialvalue=default, parent=root)
        root.destroy()
        if value is None:
            raise SystemExit(f'Cancelled: {title}')
        return int(value)
    except SystemExit:
        raise
    except Exception:
        suffix = f' [{default}]' if default is not None else ''
        raw = input(f'{prompt}{suffix}: ').strip()
        if raw == '' and default is not None:
            return int(default)
        return int(raw)


def ask_text(*, title: str, prompt: str, default: Optional[str] = None) -> str:
    try:
        import tkinter as tk
        from tkinter import simpledialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        value = simpledialog.askstring(title=title, prompt=prompt, initialvalue=default, parent=root)
        root.destroy()
        if value is None:
            raise SystemExit(f'Cancelled: {title}')
        return value
    except SystemExit:
        raise
    except Exception:
        suffix = f' [{default}]' if default else ''
        raw = input(f'{prompt}{suffix}: ').strip()
        return raw or (default or '')
