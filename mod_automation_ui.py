"""
Mod Automation Controller UI

A Tkinter desktop controller for running the seven-script mod automation workflow
without requiring users to edit Python files or type command-line paths.

The UI intentionally delegates the actual business logic to the existing scripts.
It collects/selects input files, output locations, option-period metadata, runs
preflight checks, launches each step, and streams script output to a status log.
"""

from __future__ import annotations

import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


APP_TITLE = "Mod Automation Controller"
CONFIG_FILE = "mod_automation_ui_config.json"

PROJECT_DIR = Path(__file__).resolve().parent

EXCEL_FILETYPES = [
    ("Excel files", "*.xlsx *.xlsm *.xlsb *.xls"),
    ("All files", "*.*"),
]
WORD_FILETYPES = [("Word documents", "*.docx"), ("All files", "*.*")]
JSON_FILETYPES = [("JSON files", "*.json"), ("All files", "*.*")]


@dataclass(frozen=True)
class PathField:
    key: str
    label: str
    kind: str = "file"  # file, folder, save_excel, save_word
    optional: bool = False
    help_text: str = ""


@dataclass(frozen=True)
class StepDefinition:
    number: int
    name: str
    script: str
    args: List[Tuple[str, str]]
    required_keys: List[str]
    optional_keys: List[str] = field(default_factory=list)
    output_keys: List[str] = field(default_factory=list)
    description: str = ""


FIELDS: List[PathField] = [
    PathField("output_root", "Run output folder", "folder", help_text="Parent folder for this mod package run."),
    PathField("j1_previous_file", "J.1 previous workbook", "file"),
    PathField("clin_table_file", "CLIN lookup workbook", "file"),
    PathField("pr_dir", "PR files folder", "folder"),
    PathField("source_build_file", "Source build workbook/template", "file"),
    PathField("source_j17_file", "Source J.17 workbook", "file"),
    PathField("billing_file", "Billing detail workbook", "file"),
    PathField("coversheets_dir", "Generated coversheets folder", "folder"),
    PathField("approved_coversheets_dir", "Approved coversheets folder", "folder", help_text="Use the generated coversheets folder unless approvals happen in a separate folder."),
    PathField("overview_file", "Overview workbook output", "save_excel"),
    PathField("fr_file", "F&R workbook output", "save_excel"),
    PathField("build_output_file", "Build workbook output", "save_excel"),
    PathField("j1_current_file", "J.1 current workbook output", "save_excel"),
    PathField("j17_output_file", "J.17 updated workbook output", "save_excel"),
    PathField("mfr_doc_file", "MFR walkthrough document output", "save_word"),
]

FIELD_LOOKUP = {field.key: field for field in FIELDS}

STEPS: List[StepDefinition] = [
    StepDefinition(
        1,
        "Generate coversheets",
        "01_final_coversheet_generation_script.py",
        [
            ("--j1-previous-file", "j1_previous_file"),
            ("--clin-table-file", "clin_table_file"),
            ("--pr-dir", "pr_dir"),
            ("--output-dir", "coversheets_dir"),
            ("--current-op", "current_op"),
        ],
        ["j1_previous_file", "clin_table_file", "pr_dir", "coversheets_dir", "current_op"],
        output_keys=["coversheets_dir"],
        description="Creates individual coversheet workbooks from J.1, CLIN lookup, and PR files.",
    ),
    StepDefinition(
        2,
        "Generate overview workbook",
        "02_overview_file_script.py",
        [
            ("--coversheets-dir", "coversheets_dir"),
            ("--clin-table-file", "clin_table_file"),
            ("--output-file", "overview_file"),
        ],
        ["coversheets_dir", "overview_file"],
        optional_keys=["clin_table_file"],
        output_keys=["overview_file"],
        description="Creates the overview workbook from generated coversheets. CLIN lookup is optional but recommended.",
    ),
    StepDefinition(
        3,
        "Generate F&R workbook",
        "03_F_and_R_script.py",
        [
            ("--overview-file", "overview_file"),
            ("--pr-dir", "pr_dir"),
            ("--output-file", "fr_file"),
            ("--current-option-period", "current_op"),
        ],
        ["overview_file", "pr_dir", "fr_file", "current_op"],
        output_keys=["fr_file"],
        description="Creates the F&R workbook from the overview workbook and PR folder.",
    ),
    StepDefinition(
        4,
        "Generate build workbook",
        "04_build_file_script.py",
        [
            ("--build-file", "source_build_file"),
            ("--coversheets-dir", "approved_coversheets_dir"),
            ("--pr-dir", "pr_dir"),
            ("--output-file", "build_output_file"),
        ],
        ["source_build_file", "approved_coversheets_dir", "pr_dir", "build_output_file"],
        output_keys=["build_output_file"],
        description="Creates the main build workbook from approved coversheets, PR files, and the source build workbook.",
    ),
    StepDefinition(
        5,
        "Generate current J.1 workbook",
        "05_J1_script.py",
        [
            ("--build-file", "build_output_file"),
            ("--j1-previous-file", "j1_previous_file"),
            ("--output-file", "j1_current_file"),
            ("--current-opt-pd", "current_op"),
        ],
        ["build_output_file", "j1_previous_file", "j1_current_file", "current_op"],
        output_keys=["j1_current_file"],
        description="Creates the current J.1 workbook from the build workbook and prior J.1 workbook.",
    ),
    StepDefinition(
        6,
        "Generate updated J.17 workbook",
        "06_J17_file_script.py",
        [
            ("--j17-file", "source_j17_file"),
            ("--j1-previous-file", "j1_previous_file"),
            ("--j1-current-file", "j1_current_file"),
            ("--billing-file", "billing_file"),
            ("--output-file", "j17_output_file"),
            ("--option-period", "current_op"),
            ("--current-month", "current_month"),
        ],
        ["source_j17_file", "j1_previous_file", "j1_current_file", "billing_file", "j17_output_file", "current_op"],
        optional_keys=["current_month"],
        output_keys=["j17_output_file"],
        description="Creates the updated J.17 workbook from source J.17, prior/current J.1, and billing detail files.",
    ),
    StepDefinition(
        7,
        "Generate MFR walkthrough document",
        "07_MFR_walkthrough_script.py",
        [
            ("--j1-current-file", "j1_current_file"),
            ("--output-file", "mfr_doc_file"),
            ("--option-period", "current_op"),
            ("--mod-number", "mod_number"),
        ],
        ["j1_current_file", "mfr_doc_file", "current_op", "mod_number"],
        output_keys=["mfr_doc_file"],
        description="Creates the MFR walkthrough Word document from the current J.1 workbook.",
    ),
]


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent: tk.Widget):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda event: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.canvas.bind("<Configure>", self._resize_inner)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _resize_inner(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _on_mousewheel(self, event):
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except tk.TclError:
            pass


class ModAutomationUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x820")
        self.minsize(980, 680)

        self.path_vars: Dict[str, tk.StringVar] = {field.key: tk.StringVar() for field in FIELDS}
        self.meta_vars: Dict[str, tk.StringVar] = {
            "current_op": tk.StringVar(value="5"),
            "mod_number": tk.StringVar(value="P00078"),
            "current_month": tk.StringVar(value="December"),
            "python_executable": tk.StringVar(value=sys.executable),
        }
        self.use_generated_coversheets = tk.BooleanVar(value=True)
        self.stop_requested = threading.Event()
        self.process: Optional[subprocess.Popen] = None
        self.worker_thread: Optional[threading.Thread] = None
        self.message_queue: queue.Queue = queue.Queue()
        self.step_status: Dict[int, tk.StringVar] = {
            step.number: tk.StringVar(value="Pending") for step in STEPS
        }
        self.step_buttons: Dict[int, ttk.Button] = {}
        self.running = False

        self._configure_style()
        self._build_ui()
        self._load_default_config_if_present()
        self.after(100, self._drain_queue)

    def _configure_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Subheader.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Step.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Success.TLabel", foreground="#107C10")
        style.configure("Error.TLabel", foreground="#B00020")
        style.configure("Running.TLabel", foreground="#005A9E")

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=(12, 10, 12, 6))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)

        ttk.Label(top, text=APP_TITLE, style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            top,
            text="Select the package inputs once, run preflight checks, then run each script or the complete workflow.",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=1, column=0, sticky="nsew", padx=12, pady=8)

        self.setup_tab = ttk.Frame(self.notebook)
        self.steps_tab = ttk.Frame(self.notebook)
        self.log_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.setup_tab, text="Setup")
        self.notebook.add(self.steps_tab, text="Run Workflow")
        self.notebook.add(self.log_tab, text="Log")

        self._build_setup_tab()
        self._build_steps_tab()
        self._build_log_tab()
        self._build_bottom_bar()

    def _build_setup_tab(self):
        self.setup_tab.columnconfigure(0, weight=1)
        self.setup_tab.rowconfigure(0, weight=1)
        scroller = ScrollableFrame(self.setup_tab)
        scroller.grid(row=0, column=0, sticky="nsew")
        frame = scroller.inner
        frame.columnconfigure(1, weight=1)

        general = ttk.LabelFrame(frame, text="Run metadata", padding=10)
        general.grid(row=0, column=0, columnspan=3, sticky="ew", padx=8, pady=8)
        general.columnconfigure(1, weight=1)

        self._add_text_row(general, 0, "Python executable", self.meta_vars["python_executable"], "file")
        self._add_text_row(general, 1, "Mod number", self.meta_vars["mod_number"])
        self._add_text_row(general, 2, "Current option period", self.meta_vars["current_op"])
        self._add_text_row(general, 3, "Billing current month", self.meta_vars["current_month"])
        ttk.Label(general, text="Defaults to December to preserve the original J.17 behavior; change it for the billing cycle you are processing.").grid(
            row=4, column=1, sticky="w", pady=(0, 4)
        )

        paths = ttk.LabelFrame(frame, text="Files and folders", padding=10)
        paths.grid(row=1, column=0, columnspan=3, sticky="ew", padx=8, pady=8)
        paths.columnconfigure(1, weight=1)

        for i, field in enumerate(FIELDS):
            self._add_path_row(paths, i, field)

        chk = ttk.Checkbutton(
            paths,
            text="Use generated coversheets folder as approved coversheets folder",
            variable=self.use_generated_coversheets,
            command=self._sync_approved_coversheets,
        )
        chk.grid(row=len(FIELDS), column=1, sticky="w", pady=(8, 2))

        actions = ttk.Frame(frame, padding=(8, 0, 8, 12))
        actions.grid(row=2, column=0, columnspan=3, sticky="ew")
        ttk.Button(actions, text="Set default output names", command=self.set_default_outputs).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Preflight all steps", command=lambda: self.preflight_steps(show_success=True)).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Save config", command=self.save_config_dialog).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Load config", command=self.load_config_dialog).pack(side="left", padx=(0, 8))

    def _add_text_row(self, parent, row: int, label: str, var: tk.StringVar, browse_kind: Optional[str] = None):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        if browse_kind == "file":
            ttk.Button(parent, text="Browse", command=lambda: self._browse_python(var)).grid(row=row, column=2, padx=(8, 0), pady=4)

    def _add_path_row(self, parent, row: int, field: PathField):
        ttk.Label(parent, text=field.label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        entry = ttk.Entry(parent, textvariable=self.path_vars[field.key])
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="Browse", command=lambda f=field: self.browse_path(f)).grid(row=row, column=2, padx=(8, 0), pady=4)
        if field.help_text:
            ttk.Label(parent, text=field.help_text, wraplength=280).grid(row=row, column=3, sticky="w", padx=(8, 0), pady=4)

    def _build_steps_tab(self):
        self.steps_tab.columnconfigure(0, weight=1)
        self.steps_tab.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(self.steps_tab, padding=8)
        toolbar.grid(row=0, column=0, sticky="ew")
        ttk.Button(toolbar, text="Preflight", command=lambda: self.preflight_steps(show_success=True)).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Run full workflow", command=self.run_all_steps).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Stop after current process", command=self.request_stop).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Open output folder", command=self.open_output_root).pack(side="left", padx=(0, 8))

        scroller = ScrollableFrame(self.steps_tab)
        scroller.grid(row=1, column=0, sticky="nsew")
        frame = scroller.inner
        frame.columnconfigure(0, weight=1)

        for idx, step in enumerate(STEPS):
            box = ttk.LabelFrame(frame, text=f"Step {step.number}: {step.name}", padding=10, style="Step.TLabelframe")
            box.grid(row=idx, column=0, sticky="ew", padx=8, pady=6)
            box.columnconfigure(1, weight=1)
            ttk.Label(box, text=step.description, wraplength=820).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
            ttk.Label(box, text="Status:").grid(row=1, column=0, sticky="w")
            ttk.Label(box, textvariable=self.step_status[step.number]).grid(row=1, column=1, sticky="w")
            btn = ttk.Button(box, text="Run this step", command=lambda s=step: self.run_single_step(s))
            btn.grid(row=1, column=2, sticky="e")
            self.step_buttons[step.number] = btn
            ttk.Button(box, text="Show command", command=lambda s=step: self.show_command(s)).grid(row=2, column=2, sticky="e", pady=(6, 0))

            needs = self._format_step_needs(step)
            ttk.Label(box, text=needs, wraplength=900).grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def _build_log_tab(self):
        self.log_tab.columnconfigure(0, weight=1)
        self.log_tab.rowconfigure(0, weight=1)
        self.log_text = ScrolledText(self.log_tab, wrap="word", height=22)
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self.log_text.configure(state="disabled")
        actions = ttk.Frame(self.log_tab, padding=(8, 0, 8, 8))
        actions.grid(row=1, column=0, sticky="ew")
        ttk.Button(actions, text="Clear log", command=self.clear_log).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Save log", command=self.save_log).pack(side="left")

    def _build_bottom_bar(self):
        bottom = ttk.Frame(self, padding=(12, 0, 12, 10))
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(1, weight=1)
        self.progress = ttk.Progressbar(bottom, mode="determinate", maximum=len(STEPS), value=0)
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        bottom.columnconfigure(0, weight=1)
        self.footer_status = tk.StringVar(value="Ready")
        ttk.Label(bottom, textvariable=self.footer_status).grid(row=0, column=1, sticky="e")

    def _format_step_needs(self, step: StepDefinition) -> str:
        labels = []
        for key in step.required_keys:
            if key in FIELD_LOOKUP:
                labels.append(FIELD_LOOKUP[key].label)
            elif key == "current_op":
                labels.append("current option period")
            elif key == "mod_number":
                labels.append("mod number")
            else:
                labels.append(key)
        return "Requires: " + "; ".join(labels)

    def browse_path(self, field: PathField):
        initial = self._initial_dir_for(field.key)
        current = self.path_vars[field.key].get().strip()
        selected = ""
        if field.kind == "folder":
            selected = filedialog.askdirectory(title=f"Select {field.label}", initialdir=initial)
        elif field.kind == "save_excel":
            selected = filedialog.asksaveasfilename(
                title=f"Save {field.label}",
                initialdir=initial,
                initialfile=Path(current).name if current else self._default_filename(field.key),
                filetypes=EXCEL_FILETYPES,
                defaultextension=".xlsx",
            )
        elif field.kind == "save_word":
            selected = filedialog.asksaveasfilename(
                title=f"Save {field.label}",
                initialdir=initial,
                initialfile=Path(current).name if current else self._default_filename(field.key),
                filetypes=WORD_FILETYPES,
                defaultextension=".docx",
            )
        else:
            selected = filedialog.askopenfilename(title=f"Select {field.label}", initialdir=initial, filetypes=EXCEL_FILETYPES)
        if selected:
            self.path_vars[field.key].set(str(Path(selected)))
            if field.key == "output_root":
                self.set_default_outputs(force_missing_only=True)
            if field.key == "coversheets_dir" and self.use_generated_coversheets.get():
                self._sync_approved_coversheets()

    def _browse_python(self, var: tk.StringVar):
        selected = filedialog.askopenfilename(
            title="Select Python executable",
            initialdir=str(Path(sys.executable).parent),
            filetypes=[("Python executable", "python.exe python3.exe python"), ("All files", "*.*")],
        )
        if selected:
            var.set(selected)

    def _initial_dir_for(self, key: str) -> str:
        current = self.path_vars.get(key, tk.StringVar()).get().strip()
        if current:
            p = Path(current)
            if p.is_dir():
                return str(p)
            if p.parent.exists():
                return str(p.parent)
        root = self.path_vars["output_root"].get().strip()
        if root and Path(root).exists():
            return root
        return str(PROJECT_DIR)

    def _default_filename(self, key: str) -> str:
        mod = self._safe_mod_number()
        defaults = {
            "overview_file": "overview_file.xlsx",
            "fr_file": "f_r_output.xlsx",
            "build_output_file": "build_file.xlsx",
            "j1_current_file": "j1_current_file.xlsx",
            "j17_output_file": "j17_updated_file.xlsx",
            "mfr_doc_file": f"{mod} MFR Walkthrough Output.docx",
        }
        return defaults.get(key, "output.xlsx")

    def _safe_mod_number(self) -> str:
        raw = self.meta_vars["mod_number"].get().strip() or "mod"
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", raw)
        return safe.strip("._-") or "mod"

    def _sync_approved_coversheets(self):
        if self.use_generated_coversheets.get():
            self.path_vars["approved_coversheets_dir"].set(self.path_vars["coversheets_dir"].get())

    def set_default_outputs(self, force_missing_only: bool = False):
        mod = self._safe_mod_number()
        root_text = self.path_vars["output_root"].get().strip()
        if not root_text:
            root = PROJECT_DIR / "runs" / mod
            self.path_vars["output_root"].set(str(root))
        else:
            root = Path(root_text).expanduser()

        defaults = {
            "coversheets_dir": root / "01_generated_coversheets",
            "approved_coversheets_dir": root / "01_generated_coversheets",
            "overview_file": root / "02_overview" / "overview_file.xlsx",
            "fr_file": root / "03_f_and_r" / "f_r_output.xlsx",
            "build_output_file": root / "04_build" / "build_file.xlsx",
            "j1_current_file": root / "05_j1" / "j1_current_file.xlsx",
            "j17_output_file": root / "06_j17" / "j17_updated_file.xlsx",
            "mfr_doc_file": root / "07_mfr" / f"{mod} MFR Walkthrough Output.docx",
        }
        for key, value in defaults.items():
            if force_missing_only and self.path_vars[key].get().strip():
                continue
            self.path_vars[key].set(str(value))
        if self.use_generated_coversheets.get():
            self._sync_approved_coversheets()
        self.log("Default output paths set.")

    def collect_config(self) -> dict:
        return {
            "paths": {key: var.get() for key, var in self.path_vars.items()},
            "metadata": {key: var.get() for key, var in self.meta_vars.items()},
            "use_generated_coversheets": self.use_generated_coversheets.get(),
        }

    def apply_config(self, config: dict):
        for key, value in config.get("paths", {}).items():
            if key in self.path_vars:
                self.path_vars[key].set(value or "")
        for key, value in config.get("metadata", {}).items():
            if key in self.meta_vars:
                self.meta_vars[key].set(value or "")
        self.use_generated_coversheets.set(bool(config.get("use_generated_coversheets", True)))
        if self.use_generated_coversheets.get():
            self._sync_approved_coversheets()

    def _load_default_config_if_present(self):
        config_path = PROJECT_DIR / CONFIG_FILE
        if config_path.exists():
            try:
                self.apply_config(json.loads(config_path.read_text(encoding="utf-8")))
                self.log(f"Loaded saved configuration: {config_path}")
                return
            except Exception as exc:
                self.log(f"Could not load saved configuration: {exc}")
        self.set_default_outputs(force_missing_only=True)

    def save_config_dialog(self):
        selected = filedialog.asksaveasfilename(
            title="Save workflow configuration",
            initialdir=str(PROJECT_DIR),
            initialfile=CONFIG_FILE,
            filetypes=JSON_FILETYPES,
            defaultextension=".json",
        )
        if selected:
            self.save_config(Path(selected))

    def save_config(self, path: Path):
        path.write_text(json.dumps(self.collect_config(), indent=2), encoding="utf-8")
        self.log(f"Saved configuration: {path}")

    def load_config_dialog(self):
        selected = filedialog.askopenfilename(
            title="Load workflow configuration",
            initialdir=str(PROJECT_DIR),
            filetypes=JSON_FILETYPES,
        )
        if selected:
            try:
                self.apply_config(json.loads(Path(selected).read_text(encoding="utf-8")))
                self.log(f"Loaded configuration: {selected}")
            except Exception as exc:
                messagebox.showerror(APP_TITLE, f"Could not load configuration:\n\n{exc}")

    def log(self, message: str):
        stamp = time.strftime("%H:%M:%S")
        self.message_queue.put(("log", f"[{stamp}] {message}\n"))

    def _append_log(self, text: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def save_log(self):
        selected = filedialog.asksaveasfilename(
            title="Save run log",
            initialdir=self._initial_dir_for("output_root"),
            initialfile=f"{self._safe_mod_number()}_run_log.txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            defaultextension=".txt",
        )
        if selected:
            text = self.log_text.get("1.0", "end").strip() + "\n"
            Path(selected).write_text(text, encoding="utf-8")
            self.log(f"Saved log: {selected}")

    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.message_queue.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "status":
                    step_number, value = payload
                    self.step_status[step_number].set(value)
                elif kind == "footer":
                    self.footer_status.set(payload)
                elif kind == "progress":
                    self.progress.configure(value=payload)
                elif kind == "running":
                    self._set_running_ui(payload)
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def preflight_steps(self, steps: Optional[Iterable[StepDefinition]] = None, show_success: bool = False) -> Tuple[bool, List[str]]:
        selected_steps = list(steps or STEPS)
        issues: List[str] = []
        warnings: List[str] = []

        python_exe = Path(self.meta_vars["python_executable"].get().strip() or sys.executable)
        if not python_exe.exists():
            issues.append(f"Python executable not found: {python_exe}")

        # Some path fields are outputs in one step and inputs in a later step.
        # Full-workflow preflight should allow those planned earlier outputs to be
        # missing before the run starts, but single-step preflight should require
        # them to already exist.
        planned_prior_outputs = set()

        for step in selected_steps:
            script_path = PROJECT_DIR / step.script
            if not script_path.is_file():
                issues.append(f"Step {step.number} script not found: {script_path}")

            for key in step.required_keys:
                value = self._value_for_key(key)
                label = self._label_for_key(key)
                if not value:
                    issues.append(f"Step {step.number} missing required value: {label}")
                    continue
                if key in FIELD_LOOKUP:
                    field = FIELD_LOOKUP[key]
                    p = Path(value).expanduser()
                    is_output_for_this_step = key in step.output_keys
                    will_be_created_by_prior_step = key in planned_prior_outputs

                    if is_output_for_this_step:
                        if field.kind == "folder":
                            parent = p.parent
                            if parent and not parent.exists():
                                warnings.append(f"Step {step.number} will create output folder under: {parent}")
                        elif field.kind.startswith("save_"):
                            parent = p.parent
                            if parent and not parent.exists():
                                warnings.append(f"Step {step.number} will create output folder: {parent}")
                    elif will_be_created_by_prior_step:
                        warnings.append(f"Step {step.number} will use prior workflow output: {label}: {p}")
                    elif field.kind == "file" and not p.is_file():
                        issues.append(f"Step {step.number} input file not found: {label}: {p}")
                    elif field.kind == "folder" and not p.is_dir():
                        issues.append(f"Step {step.number} input folder not found: {label}: {p}")
                    elif field.kind.startswith("save_") and not p.is_file():
                        issues.append(f"Step {step.number} input file not found: {label}: {p}")

            for key in step.optional_keys:
                value = self._value_for_key(key)
                if key in FIELD_LOOKUP and value:
                    p = Path(value).expanduser()
                    field = FIELD_LOOKUP[key]
                    if field.kind == "file" and not p.is_file():
                        warnings.append(f"Step {step.number} optional file not found and will be skipped: {self._label_for_key(key)}: {p}")

            planned_prior_outputs.update(step.output_keys)

        if issues:
            message = "Preflight found issues:\n\n" + "\n".join(f"- {item}" for item in issues)
            if warnings:
                message += "\n\nWarnings:\n" + "\n".join(f"- {item}" for item in warnings)
            messagebox.showerror(APP_TITLE, message)
            self.log("Preflight failed.")
            for item in issues:
                self.log(f"PRECHECK ERROR: {item}")
            for item in warnings:
                self.log(f"PRECHECK WARNING: {item}")
            return False, issues

        for item in warnings:
            self.log(f"PRECHECK WARNING: {item}")
        if show_success:
            messagebox.showinfo(APP_TITLE, "Preflight passed. Required inputs are present and output paths are ready to be created as needed.")
        self.log("Preflight passed.")
        return True, []

    def _label_for_key(self, key: str) -> str:
        if key in FIELD_LOOKUP:
            return FIELD_LOOKUP[key].label
        labels = {
            "current_op": "current option period",
            "mod_number": "mod number",
            "current_month": "current billing month",
        }
        return labels.get(key, key)

    def _value_for_key(self, key: str) -> str:
        if key in self.path_vars:
            return self.path_vars[key].get().strip()
        if key in self.meta_vars:
            return self.meta_vars[key].get().strip()
        return ""

    def build_command(self, step: StepDefinition) -> List[str]:
        python_exe = self.meta_vars["python_executable"].get().strip() or sys.executable
        command = [python_exe, str(PROJECT_DIR / step.script)]
        for flag, key in step.args:
            value = self._value_for_key(key)
            if not value and key in step.optional_keys:
                continue
            if value:
                command.extend([flag, value])
        return command

    def show_command(self, step: StepDefinition):
        cmd = self.build_command(step)
        rendered = " ".join(self._quote_part(part) for part in cmd)
        messagebox.showinfo(f"Step {step.number} command", rendered)

    @staticmethod
    def _quote_part(value: str) -> str:
        if re.search(r"\s", value):
            return f'"{value}"'
        return value

    def run_single_step(self, step: StepDefinition):
        if self.running:
            messagebox.showwarning(APP_TITLE, "A workflow process is already running.")
            return
        ok, _ = self.preflight_steps([step])
        if not ok:
            return
        self.stop_requested.clear()
        self.worker_thread = threading.Thread(target=self._run_steps_worker, args=([step],), daemon=True)
        self.worker_thread.start()

    def run_all_steps(self):
        if self.running:
            messagebox.showwarning(APP_TITLE, "A workflow process is already running.")
            return
        ok, _ = self.preflight_steps(STEPS)
        if not ok:
            return
        self.stop_requested.clear()
        self.worker_thread = threading.Thread(target=self._run_steps_worker, args=(STEPS,), daemon=True)
        self.worker_thread.start()

    def request_stop(self):
        self.stop_requested.set()
        self.log("Stop requested. The current script will be terminated if it is still running.")
        if self.process and self.process.poll() is None:
            try:
                self.process.terminate()
            except Exception as exc:
                self.log(f"Could not terminate process cleanly: {exc}")

    def _run_steps_worker(self, steps: Iterable[StepDefinition]):
        steps = list(steps)
        self.message_queue.put(("running", True))
        self.message_queue.put(("footer", "Running"))
        try:
            for step in steps:
                if self.stop_requested.is_set():
                    self.message_queue.put(("status", (step.number, "Skipped")))
                    continue
                self.message_queue.put(("status", (step.number, "Running")))
                self.message_queue.put(("footer", f"Running Step {step.number}: {step.name}"))
                self.log(f"----- Step {step.number}: {step.name} -----")
                cmd = self.build_command(step)
                self.log("Command: " + " ".join(self._quote_part(part) for part in cmd))
                self._ensure_output_locations(step)

                try:
                    self.process = subprocess.Popen(
                        cmd,
                        cwd=str(PROJECT_DIR),
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        bufsize=1,
                    )
                    assert self.process.stdout is not None
                    for line in self.process.stdout:
                        self.message_queue.put(("log", line))
                    return_code = self.process.wait()
                except FileNotFoundError as exc:
                    self.log(f"ERROR: {exc}")
                    return_code = 127
                except Exception as exc:
                    self.log(f"ERROR: {exc}")
                    return_code = 1
                finally:
                    self.process = None

                if return_code == 0:
                    self.message_queue.put(("status", (step.number, "Complete")))
                    self.log(f"Step {step.number} complete.")
                    self.message_queue.put(("progress", step.number))
                else:
                    self.message_queue.put(("status", (step.number, "Failed")))
                    self.log(f"Step {step.number} failed with exit code {return_code}.")
                    self.message_queue.put(("footer", f"Stopped at Step {step.number}"))
                    return
            self.message_queue.put(("footer", "Workflow complete"))
            self.log("Workflow complete.")
            try:
                self.save_config(PROJECT_DIR / CONFIG_FILE)
            except Exception as exc:
                self.log(f"Could not save default configuration: {exc}")
        finally:
            self.message_queue.put(("running", False))

    def _ensure_output_locations(self, step: StepDefinition):
        for key in step.output_keys:
            value = self._value_for_key(key)
            if not value or key not in FIELD_LOOKUP:
                continue
            field = FIELD_LOOKUP[key]
            path = Path(value).expanduser()
            if field.kind == "folder":
                path.mkdir(parents=True, exist_ok=True)
            elif field.kind.startswith("save_"):
                path.parent.mkdir(parents=True, exist_ok=True)

    def _set_running_ui(self, running: bool):
        self.running = running
        state = "disabled" if running else "normal"
        for btn in self.step_buttons.values():
            btn.configure(state=state)
        if running:
            self.notebook.select(self.log_tab)
        else:
            self.footer_status.set(self.footer_status.get() if self.footer_status.get() != "Running" else "Ready")

    def open_output_root(self):
        root = self.path_vars["output_root"].get().strip()
        if not root:
            messagebox.showwarning(APP_TITLE, "No output folder is set.")
            return
        path = Path(root).expanduser()
        path.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not open output folder:\n\n{exc}")


def main():
    app = ModAutomationUI()
    app.mainloop()


if __name__ == "__main__":
    main()
