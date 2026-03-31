from __future__ import annotations

import argparse
import importlib.util
import json
import os
import tempfile
import sys
import traceback
import types
from pathlib import Path
from typing import Any, Dict, List

from ui_workflow import STEP_MAP




def ensure_runtime_environment() -> None:
    home = os.environ.get("HOME") or os.environ.get("USERPROFILE")
    if not home:
        drive = os.environ.get("HOMEDRIVE", "")
        path = os.environ.get("HOMEPATH", "")
        candidate = f"{drive}{path}".strip()
        if candidate:
            home = candidate
    if not home:
        home = tempfile.gettempdir()
    os.environ.setdefault("HOME", home)
    os.environ.setdefault("USERPROFILE", home)
    mpl_dir = Path(home) / ".matplotlib"
    try:
        mpl_dir.mkdir(parents=True, exist_ok=True)
        os.environ.setdefault("MPLCONFIGDIR", str(mpl_dir))
    except Exception:
        os.environ.setdefault("MPLCONFIGDIR", tempfile.gettempdir())


class ResponseQueue:
    def __init__(self, responses: List[Dict[str, Any]]):
        self._responses = list(responses)

    def pop(self, expected_type: str) -> Any:
        if not self._responses:
            raise RuntimeError(f"No more configured responses available for {expected_type}.")
        item = self._responses.pop(0)
        actual_type = item.get("type")
        if actual_type != expected_type:
            raise RuntimeError(f"Expected response type {expected_type}, got {actual_type}.")
        return item.get("value")


class WorkflowProxyFactory:
    def __init__(self, queue: ResponseQueue):
        self.queue = queue

    def build_module(self) -> types.ModuleType:
        mod = types.ModuleType("workflow_io")

        def choose_file(**_: Any) -> Path:
            return Path(self.queue.pop("file"))

        def choose_files(**_: Any) -> List[Path]:
            return [Path(p) for p in self.queue.pop("files")]

        def choose_directory(**_: Any) -> Path:
            return Path(self.queue.pop("directory"))

        def choose_save_file(**_: Any) -> Path:
            return Path(self.queue.pop("save_file"))

        def ask_integer(**_: Any) -> int:
            return int(self.queue.pop("integer"))

        def ask_text(**_: Any) -> str:
            return str(self.queue.pop("text"))

        mod.choose_file = choose_file
        mod.choose_files = choose_files
        mod.choose_directory = choose_directory
        mod.choose_save_file = choose_save_file
        mod.ask_integer = ask_integer
        mod.ask_text = ask_text
        return mod


def load_module(script_path: Path, proxy_module: types.ModuleType):
    original = sys.modules.get("workflow_io")
    sys.modules["workflow_io"] = proxy_module
    module_name = f"workflow_step_{script_path.stem}"
    try:
        spec = importlib.util.spec_from_file_location(module_name, script_path)
        if spec is None or spec.loader is None:
            raise RuntimeError(f"Unable to load script: {script_path}")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if original is not None:
            sys.modules["workflow_io"] = original
        else:
            sys.modules.pop("workflow_io", None)


def run_step(step_id: str, config_path: Path) -> int:
    step = STEP_MAP[step_id]
    config = json.loads(config_path.read_text(encoding="utf-8"))
    queue = ResponseQueue(config["responses"])
    proxy = WorkflowProxyFactory(queue).build_module()
    script_path = Path(__file__).with_name(step.script_name)
    module = load_module(script_path, proxy)

    if step.entry_mode == "main":
        getattr(module, step.entry_function)()
    elif step.entry_mode == "configure_then_call":
        module.configure_runtime()
        getattr(module, step.entry_function)()
    else:
        raise RuntimeError(f"Unsupported entry mode: {step.entry_mode}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Run a workflow step with injected UI responses.")
    parser.add_argument("--step", required=True)
    parser.add_argument("--config", required=True)
    args = parser.parse_args()

    ensure_runtime_environment()
    try:
        return run_step(args.step, Path(args.config))
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 1
        if code not in (0, None):
            print(f"\nStep cancelled or exited: {exc}")
        return int(code or 0)
    except Exception:
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
