from __future__ import annotations

import json
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from PySide6.QtCore import QProcess, QSize, Qt, Signal, QTimer, QProcessEnvironment
from PySide6.QtGui import QAction, QDesktopServices, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QProgressBar,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QFileDialog,
    QHeaderView,
)

from ui_workflow import FieldDef, STEP_MAP, STEPS, StepDef

APP_STATE_FILE = Path(__file__).with_name("ui_state.json")
APP_TITLE = "Contract Mod Automation"


class AppState:
    def __init__(self) -> None:
        self.data = {
            "steps": {},
            "history": [],
            "settings": {"auto_open_output_folder": False},
        }
        self.load()

    def load(self) -> None:
        if APP_STATE_FILE.exists():
            try:
                self.data = json.loads(APP_STATE_FILE.read_text(encoding="utf-8"))
            except Exception:
                pass

    def save(self) -> None:
        APP_STATE_FILE.write_text(json.dumps(self.data, indent=2), encoding="utf-8")

    def step_data(self, step_id: str) -> dict:
        self.data.setdefault("steps", {})
        self.data["steps"].setdefault(step_id, {})
        return self.data["steps"][step_id]

    def add_history(self, entry: dict) -> None:
        history = self.data.setdefault("history", [])
        history.insert(0, entry)
        del history[100:]
        self.save()


def make_button(text: str, role: str = "secondary", tooltip: str = "") -> QPushButton:
    button = QPushButton(text)
    button.setProperty("role", role)
    if tooltip:
        button.setToolTip(tooltip)
    return button


class StepCard(QFrame):
    clicked = Signal(str)

    def __init__(self, step: StepDef):
        super().__init__()
        self.step = step
        self.setObjectName("StepCard")
        self.setFrameShape(QFrame.StyledPanel)
        self.setCursor(Qt.PointingHandCursor)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        self.title_label = QLabel(step.title)
        self.title_label.setStyleSheet("font-weight: 600; font-size: 14px;")
        self.status_label = QLabel("Not configured")
        self.status_label.setObjectName("stepStatus")
        self.detail_label = QLabel(step.description)
        self.detail_label.setWordWrap(True)
        self.detail_label.setStyleSheet("color: #4b5563;")
        layout.addWidget(self.title_label)
        layout.addWidget(self.status_label)
        layout.addWidget(self.detail_label)

    def mousePressEvent(self, event):
        self.clicked.emit(self.step.step_id)
        return super().mousePressEvent(event)

    def update_status(self, text: str) -> None:
        self.status_label.setText(text)


class OverviewPage(QWidget):
    navigate_to_step = Signal(str)

    def __init__(self, state: AppState):
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        header = QLabel("Workflow Overview")
        header.setStyleSheet("font-size: 20px; font-weight: 700;")
        layout.addWidget(header)
        sub = QLabel("Run the six-step contract modification workflow from one interface. Configure only the inputs for the step you are on, and reuse prior outputs where available.")
        sub.setWordWrap(True)
        sub.setStyleSheet("color: #555;")
        layout.addWidget(sub)

        self.summary_label = QLabel()
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(12)
        self.cards: Dict[str, StepCard] = {}
        for idx, step in enumerate(STEPS):
            card = StepCard(step)
            card.clicked.connect(self.navigate_to_step)
            self.cards[step.step_id] = card
            grid.addWidget(card, idx // 2, idx % 2)
        layout.addLayout(grid)
        layout.addStretch(1)

    def refresh(self) -> None:
        completed = 0
        for step in STEPS:
            step_state = self.state.step_data(step.step_id)
            status = step_state.get("status", "not_configured")
            last_run = step_state.get("last_run")
            status_text = {
                "not_configured": "Not configured",
                "ready": "Ready to run",
                "running": "Running",
                "success": f"Completed{f' • {last_run}' if last_run else ''}",
                "failed": f"Last run failed{f' • {last_run}' if last_run else ''}",
            }.get(status, status)
            self.cards[step.step_id].update_status(status_text)
            if status == "success":
                completed += 1
        self.summary_label.setText(f"Completed steps: {completed} of {len(STEPS)}")


class HistoryPage(QWidget):
    def __init__(self, state: AppState):
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        header = QLabel("Run History")
        header.setStyleSheet("font-size: 20px; font-weight: 700;")
        layout.addWidget(header)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Time", "Step", "Result", "Exit Code", "Output"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table)

    def refresh(self) -> None:
        history = self.state.data.get("history", [])
        self.table.setRowCount(len(history))
        for row, entry in enumerate(history):
            values = [
                entry.get("timestamp", ""),
                entry.get("step_title", ""),
                entry.get("result", ""),
                str(entry.get("exit_code", "")),
                entry.get("latest_output", ""),
            ]
            for col, value in enumerate(values):
                self.table.setItem(row, col, QTableWidgetItem(value))
        self.table.resizeColumnsToContents()


class SettingsPage(QWidget):
    settings_changed = Signal()

    def __init__(self, state: AppState):
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        header = QLabel("Settings")
        header.setStyleSheet("font-size: 20px; font-weight: 700;")
        layout.addWidget(header)
        box = QGroupBox("General")
        form = QFormLayout(box)
        self.auto_open_checkbox = QCheckBox("Open the output folder automatically after a successful run")
        self.auto_open_checkbox.setToolTip("When enabled, the folder containing the generated file opens after a successful run.")
        self.auto_open_checkbox.setChecked(bool(self.state.data.get("settings", {}).get("auto_open_output_folder", False)))
        self.auto_open_checkbox.toggled.connect(self.save_settings)
        form.addRow(self.auto_open_checkbox)
        layout.addWidget(box)
        layout.addStretch(1)

    def save_settings(self) -> None:
        self.state.data.setdefault("settings", {})["auto_open_output_folder"] = self.auto_open_checkbox.isChecked()
        self.state.save()
        self.settings_changed.emit()


class HelpPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        header = QLabel("Setup and Run Guide")
        header.setStyleSheet("font-size: 20px; font-weight: 700;")
        layout.addWidget(header)
        text = QTextEdit()
        text.setReadOnly(True)
        text.setPlainText(
            "Use this UI to run the six workflow steps without a terminal.\n\n"
            "How to use:\n"
            "1. Open a step from the left rail.\n"
            "2. Fill in the required inputs for that step.\n"
            "3. Use 'Use Previous Output' when a step depends on an earlier generated file or folder.\n"
            "4. Click Run Step.\n"
            "5. Review the log and run summary in the same step page.\n\n"
            "Important note for Step 5:\n"
            "The source J.1 workbook should be a true .xlsx or .xlsm workbook. The current Script 05 workflow does not safely update .xlsb files."
        )
        layout.addWidget(text)


class StepPage(QWidget):
    run_requested = Signal(str, dict)
    navigate_to_step = Signal(str)

    def __init__(self, step: StepDef, state: AppState):
        super().__init__()
        self.step = step
        self.state = state
        self.editors: Dict[str, QWidget] = {}
        self.output_open_buttons: Dict[str, QPushButton] = {}
        self.latest_output_path: Optional[str] = None

        root = QVBoxLayout(self)
        title = QLabel(step.title)
        title.setStyleSheet("font-size: 20px; font-weight: 700;")
        root.addWidget(title)
        desc = QLabel(step.description)
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #555;")
        root.addWidget(desc)

        if step.notes:
            notes = QLabel("\n".join(step.notes))
            notes.setWordWrap(True)
            notes.setStyleSheet("color: #8a5a00; background: #fff6db; padding: 8px; border: 1px solid #f0d58a;")
            root.addWidget(notes)

        prereq_box = QGroupBox("Prerequisites")
        prereq_layout = QVBoxLayout(prereq_box)
        self.prereq_label = QLabel()
        self.prereq_label.setWordWrap(True)
        prereq_layout.addWidget(self.prereq_label)
        root.addWidget(prereq_box)

        form_box = QGroupBox("Inputs and Outputs")
        form_layout = QFormLayout(form_box)
        form_layout.setLabelAlignment(Qt.AlignRight | Qt.AlignTop)
        for field_def in step.fields:
            editor_row = self._build_field_row(field_def)
            form_layout.addRow(field_def.label + (" *" if field_def.required else ""), editor_row)
        root.addWidget(form_box)

        controls = QHBoxLayout()
        self.validate_button = make_button("Validate", tooltip="Check required fields and upstream dependencies before running.")
        self.validate_button.clicked.connect(self.validate_inputs)
        controls.addWidget(self.validate_button)
        self.run_button = make_button("Run Step", role="primary", tooltip="Run the current workflow step with the values shown above.")
        self.run_button.clicked.connect(self._emit_run)
        controls.addWidget(self.run_button)
        self.reset_button = make_button("Reset Step", tooltip="Clear the current step values and validation state.")
        self.reset_button.clicked.connect(self.reset_fields)
        controls.addWidget(self.reset_button)
        controls.addStretch(1)
        root.addLayout(controls)

        running_row = QHBoxLayout()
        self.running_indicator = QLabel("Idle")
        self.running_indicator.setObjectName("runningIndicator")
        self.running_indicator.setStyleSheet("background: #f3f4f6; color: #374151; padding: 6px 10px; border: 1px solid #d1d5db; border-radius: 14px; font-weight: 600;")
        running_row.addWidget(self.running_indicator)
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(14)
        running_row.addWidget(self.progress_bar, 1)
        root.addLayout(running_row)

        self.spinner_timer = QTimer(self)
        self.spinner_timer.setInterval(180)
        self._spinner_frames = ["Running   ", "Running.  ", "Running.. ", "Running..."]
        self._spinner_index = 0
        self.spinner_timer.timeout.connect(self._advance_spinner)

        status_box = QGroupBox("Validation and Run Summary")
        status_layout = QVBoxLayout(status_box)
        self.validation_label = QLabel("No validation results yet.")
        self.validation_label.setObjectName("validationLabel")
        self.validation_label.setWordWrap(True)
        self.validation_label.setStyleSheet("padding: 6px;")
        status_layout.addWidget(self.validation_label)
        self.summary_label = QLabel("Step has not been run yet.")
        self.summary_label.setWordWrap(True)
        status_layout.addWidget(self.summary_label)
        root.addWidget(status_box)

        log_box = QGroupBox("Step Log")
        log_layout = QVBoxLayout(log_box)
        self.log_edit = QPlainTextEdit()
        self.log_edit.setPlaceholderText("Run output and processing messages will appear here.")
        self.log_edit.setReadOnly(True)
        self.log_edit.setMaximumBlockCount(5000)
        log_layout.addWidget(self.log_edit)
        root.addWidget(log_box, 1)

        self.load_from_state()
        self.refresh_prereq_status()

    def _build_field_row(self, field_def: FieldDef) -> QWidget:
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        if field_def.kind == "integer":
            editor = QSpinBox()
            editor.setMaximum(999)
            editor.setValue(int(field_def.placeholder or 0))
            editor.valueChanged.connect(self._save_current_state)
            layout.addWidget(editor, 1)
        else:
            editor = QLineEdit()
            editor.setClearButtonEnabled(True)
            editor.setPlaceholderText(field_def.placeholder)
            editor.textChanged.connect(self._save_current_state)
            layout.addWidget(editor, 1)
            if field_def.kind in {"file", "directory", "save_file"}:
                browse = make_button("Browse", tooltip="Browse for this path.")
                browse.clicked.connect(lambda _=False, f=field_def: self.browse_for_field(f))
                layout.addWidget(browse)
            if field_def.source_step_id and field_def.source_field_key:
                use_prev = make_button("Use Previous Output", tooltip="Populate this field from the saved output of the linked upstream step.")
                use_prev.clicked.connect(lambda _=False, f=field_def: self.use_previous_output(f))
                layout.addWidget(use_prev)
            open_btn = make_button("Open", tooltip="Open this file or folder if the path exists.")
            open_btn.clicked.connect(lambda _=False, key=field_def.key: self.open_path(key))
            layout.addWidget(open_btn)
            if field_def.is_output:
                self.output_open_buttons[field_def.key] = open_btn
            copy_btn = make_button("Copy", tooltip="Copy this path to the clipboard.")
            copy_btn.clicked.connect(lambda _=False, key=field_def.key: self.copy_path(key))
            layout.addWidget(copy_btn)
        self.editors[field_def.key] = editor
        return container

    def field_value(self, field_key: str):
        editor = self.editors[field_key]
        if isinstance(editor, QSpinBox):
            return editor.value()
        return editor.text().strip()

    def set_field_value(self, field_key: str, value):
        editor = self.editors[field_key]
        if isinstance(editor, QSpinBox):
            editor.setValue(int(value or 0))
        else:
            editor.setText(str(value or ""))

    def browse_for_field(self, field_def: FieldDef) -> None:
        current = self.field_value(field_def.key)
        if field_def.kind == "file":
            path, _ = QFileDialog.getOpenFileName(self, field_def.label, current or str(Path.home()), field_def.file_filter)
        elif field_def.kind == "directory":
            path = QFileDialog.getExistingDirectory(self, field_def.label, current or str(Path.home()))
        elif field_def.kind == "save_file":
            path, _ = QFileDialog.getSaveFileName(self, field_def.label, current or field_def.default_name or str(Path.home()), field_def.file_filter)
        else:
            return
        if path:
            self.set_field_value(field_def.key, path)

    def use_previous_output(self, field_def: FieldDef) -> None:
        if not field_def.source_step_id or not field_def.source_field_key:
            return
        source_state = self.state.step_data(field_def.source_step_id)
        source_values = source_state.get("values", {})
        value = source_values.get(field_def.source_field_key, "")
        if not value:
            QMessageBox.information(self, "Previous output unavailable", "No saved output is available from the upstream step yet.")
            return
        self.set_field_value(field_def.key, value)

    def open_path(self, field_key: str) -> None:
        path = self.field_value(field_key)
        if not path:
            return
        resolved = Path(path)
        if resolved.exists():
            QDesktopServices.openUrl(resolved.as_uri())
            return
        parent = resolved.parent
        if parent.exists():
            QDesktopServices.openUrl(parent.as_uri())

    def copy_path(self, field_key: str) -> None:
        path = self.field_value(field_key)
        if path:
            QApplication.clipboard().setText(str(path))

    def build_values(self) -> dict:
        return {field.key: self.field_value(field.key) for field in self.step.fields}

    def build_responses(self) -> list:
        responses = []
        for field in self.step.fields:
            value = self.field_value(field.key)
            response_type = {
                "file": "file",
                "directory": "directory",
                "save_file": "save_file",
                "integer": "integer",
                "text": "text",
            }[field.kind]
            responses.append({"type": response_type, "value": value})
        return responses

    def validate_inputs(self) -> bool:
        messages: List[str] = []
        warnings: List[str] = []
        for field in self.step.fields:
            value = self.field_value(field.key)
            if field.required and (value == "" or value is None):
                messages.append(f"{field.label} is required.")
                continue
            if field.kind == "file" and value and not Path(value).is_file():
                messages.append(f"{field.label} does not point to an existing file.")
            elif field.kind == "directory" and value and not Path(value).is_dir():
                messages.append(f"{field.label} does not point to an existing folder.")
            elif field.kind == "save_file" and value:
                parent = Path(value).parent
                if not parent.exists():
                    messages.append(f"Output folder for {field.label} does not exist.")
            if field.validator and value:
                warning = field.validator(str(value))
                if warning:
                    warnings.append(warning)
        prereq_warning = self.prerequisite_warning()
        if prereq_warning:
            warnings.append(prereq_warning)

        if messages:
            self.validation_label.setStyleSheet("background: #ffe7e7; color: #8b1e1e; padding: 6px; border: 1px solid #e3a9a9;")
            self.validation_label.setText("\n".join(messages + (["", *warnings] if warnings else [])))
            return False

        if warnings:
            self.validation_label.setStyleSheet("background: #fff6db; color: #8a5a00; padding: 6px; border: 1px solid #f0d58a;")
            self.validation_label.setText("Validation passed with notes:\n" + "\n".join(warnings))
        else:
            self.validation_label.setStyleSheet("background: #e8f6ea; color: #205b2d; padding: 6px; border: 1px solid #9cccab;")
            self.validation_label.setText("All required fields are ready.")
        return True

    def prerequisite_warning(self) -> Optional[str]:
        stale_steps = []
        for prereq in self.step.prerequisites:
            prereq_state = self.state.step_data(prereq)
            if prereq_state.get("status") != "success":
                stale_steps.append(STEP_MAP[prereq].short_title)
        if stale_steps:
            return "Upstream steps not completed yet: " + ", ".join(stale_steps)
        return None

    def refresh_prereq_status(self) -> None:
        lines: List[str] = []

        if self.step.prerequisites:
            lines.append("Upstream workflow requirements:")
            for prereq in self.step.prerequisites:
                prereq_state = self.state.step_data(prereq)
                status = prereq_state.get("status", "not_configured")
                label = {
                    "success": "Completed",
                    "running": "Running",
                    "failed": "Failed",
                    "ready": "Configured",
                    "not_configured": "Not configured",
                }.get(status, status)
                outputs = STEP_MAP[prereq].outputs
                output_text = ""
                if outputs:
                    source_values = prereq_state.get("values", {})
                    linked = [str(source_values.get(key, "")).strip() for key in outputs if str(source_values.get(key, "")).strip()]
                    if linked:
                        output_text = f" — output: {linked[0]}"
                lines.append(f"• {STEP_MAP[prereq].short_title}: {label}{output_text}")
        else:
            lines.append("Upstream workflow requirements:")
            lines.append("• None")

        lines.append("")
        lines.append("Required inputs for this step:")
        for field in self.step.fields:
            if not field.required:
                continue
            source_note = ""
            if field.source_step_id and field.source_field_key:
                source_note = f" (can use previous output from {STEP_MAP[field.source_step_id].short_title})"
            kind_label = {
                "file": "File",
                "directory": "Folder",
                "save_file": "Save location",
                "integer": "Number",
                "text": "Text",
            }.get(field.kind, field.kind.title())
            lines.append(f"• {field.label}: {kind_label}{source_note}")

        self.prereq_label.setText("\n".join(lines))

    def reset_fields(self) -> None:
        for field in self.step.fields:
            if field.kind == "integer":
                self.set_field_value(field.key, field.placeholder or 0)
            else:
                self.set_field_value(field.key, "")
        self.log_edit.clear()
        self.summary_label.setText("Step has not been run yet.")
        self.validation_label.setText("No validation results yet.")
        self.validation_label.setStyleSheet("background: #f3f4f6; color: #111827; padding: 8px; border: 1px solid #d1d5db; border-radius: 6px;")
        self._save_current_state()

    def append_log(self, text: str) -> None:
        self.log_edit.moveCursor(QTextCursor.End)
        self.log_edit.insertPlainText(text)
        self.log_edit.moveCursor(QTextCursor.End)

    def set_running(self, running: bool) -> None:
        self.run_button.setEnabled(not running)
        self.validate_button.setEnabled(not running)
        self.reset_button.setEnabled(not running)
        if running:
            self.summary_label.setText("Running...")
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            self._spinner_index = 0
            self.running_indicator.setText(self._spinner_frames[self._spinner_index])
            self.running_indicator.setStyleSheet("background: #dbeafe; color: #1d4ed8; padding: 6px 10px; border: 1px solid #93c5fd; border-radius: 14px; font-weight: 700;")
            self.spinner_timer.start()
        else:
            self.spinner_timer.stop()
            self.progress_bar.setVisible(False)
            self.progress_bar.setRange(0, 1)
            self.progress_bar.setValue(0)
            self.running_indicator.setText("Idle")
            self.running_indicator.setStyleSheet("background: #f3f4f6; color: #374151; padding: 6px 10px; border: 1px solid #d1d5db; border-radius: 14px; font-weight: 600;")

    def _advance_spinner(self) -> None:
        self._spinner_index = (self._spinner_index + 1) % len(self._spinner_frames)
        self.running_indicator.setText(self._spinner_frames[self._spinner_index])

    def _emit_run(self) -> None:
        if not self.validate_inputs():
            return
        values = self.build_values()
        payload = {
            "step_id": self.step.step_id,
            "values": values,
            "responses": self.build_responses(),
        }
        self.run_requested.emit(self.step.step_id, payload)

    def load_from_state(self) -> None:
        step_state = self.state.step_data(self.step.step_id)
        values = step_state.get("values", {})
        for field in self.step.fields:
            if field.key in values:
                self.set_field_value(field.key, values[field.key])
        if step_state.get("last_log"):
            self.log_edit.setPlainText(step_state["last_log"])
        if step_state.get("last_summary"):
            self.summary_label.setText(step_state["last_summary"])

    def save_to_state(self, extra: Optional[dict] = None) -> None:
        step_state = self.state.step_data(self.step.step_id)
        step_state["values"] = self.build_values()
        if extra:
            step_state.update(extra)
        self.state.save()

    def _save_current_state(self, *args) -> None:
        self.save_to_state()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.state = AppState()
        self.current_process: Optional[QProcess] = None
        self.current_temp_config: Optional[Path] = None
        self.current_step_id: Optional[str] = None
        self.current_log_buffer: str = ""
        self.setWindowTitle(APP_TITLE)
        self.resize(1500, 900)
        self._build_ui()
        self.apply_styles()
        self.refresh_all()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)

        splitter = QSplitter()
        splitter.setChildrenCollapsible(False)
        main_layout.addWidget(splitter)

        self.nav_list = QListWidget()
        self.nav_list.setFixedWidth(250)
        self.nav_list.currentRowChanged.connect(self.switch_page)
        splitter.addWidget(self.nav_list)

        self.stack = QStackedWidget()
        splitter.addWidget(self.stack)

        self.context_panel = QFrame()
        self.context_panel.setObjectName("context_panel")
        self.context_panel.setFixedWidth(320)
        ctx_layout = QVBoxLayout(self.context_panel)
        ctx_layout.setContentsMargins(12, 12, 12, 12)
        ctx_title = QLabel("Current Context")
        ctx_title.setStyleSheet("font-size: 16px; font-weight: 700;")
        ctx_layout.addWidget(ctx_title)
        self.context_step = QLabel("Home")
        self.context_step.setStyleSheet("font-weight: 600;")
        ctx_layout.addWidget(self.context_step)
        self.context_desc = QLabel("")
        self.context_desc.setWordWrap(True)
        ctx_layout.addWidget(self.context_desc)
        self.context_status = QLabel("")
        self.context_status.setWordWrap(True)
        ctx_layout.addWidget(self.context_status)
        self.context_output = QLabel("")
        self.context_output.setWordWrap(True)
        ctx_layout.addWidget(self.context_output)
        self.open_output_btn = make_button("Open Latest Output", tooltip="Open the latest generated file or folder for the current step.")
        self.open_output_btn.clicked.connect(self.open_latest_output)
        ctx_layout.addWidget(self.open_output_btn)
        self.context_log = QPlainTextEdit()
        self.context_log.setPlaceholderText("Current step log preview.")
        self.context_log.setReadOnly(True)
        ctx_layout.addWidget(self.context_log, 1)
        splitter.addWidget(self.context_panel)
        splitter.setSizes([240, 940, 320])

        self.overview_page = OverviewPage(self.state)
        self.overview_page.navigate_to_step.connect(self.focus_step)
        self.history_page = HistoryPage(self.state)
        self.settings_page = SettingsPage(self.state)
        self.help_page = HelpPage()
        self.step_pages: Dict[str, StepPage] = {}

        self._add_nav_item("Overview")
        self.stack.addWidget(self.overview_page)
        for step in STEPS:
            self._add_nav_item(step.title)
            page = StepPage(step, self.state)
            page.run_requested.connect(self.run_step)
            self.stack.addWidget(page)
            self.step_pages[step.step_id] = page
        self._add_nav_item("Run History")
        self.stack.addWidget(self.history_page)
        self._add_nav_item("Settings")
        self.stack.addWidget(self.settings_page)
        self._add_nav_item("Help / Run Guide")
        self.stack.addWidget(self.help_page)

        self.nav_list.setCurrentRow(0)

        toolbar = self.addToolBar("Main")
        self.statusBar().showMessage("Ready")
        toolbar.setMovable(False)
        refresh_action = QAction("Refresh", self)
        refresh_action.triggered.connect(self.refresh_all)
        toolbar.addAction(refresh_action)

    def _add_nav_item(self, text: str) -> None:
        item = QListWidgetItem(text)
        item.setSizeHint(QSize(240, 42))
        self.nav_list.addItem(item)

    def apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #eef2f7; color: #111827; }
            QToolBar { background: #ffffff; border-bottom: 1px solid #d1d5db; spacing: 6px; padding: 6px; }
            QStatusBar { background: #ffffff; color: #111827; border-top: 1px solid #d1d5db; }
            QListWidget { background: #111827; color: #f9fafb; border: none; padding: 8px 0; outline: none; }
            QListWidget::item { padding: 10px 14px; border-left: 4px solid transparent; }
            QListWidget::item:selected { background: #1d4ed8; color: #ffffff; border-left: 4px solid #93c5fd; }
            QListWidget::item:hover:!selected { background: #1f2937; }
            QFrame#StepCard, QFrame#context_panel, QGroupBox, QTextEdit, QPlainTextEdit, QTableWidget { background: #ffffff; color: #111827; }
            QGroupBox { font-weight: 600; border: 1px solid #d0d7de; border-radius: 10px; margin-top: 12px; padding-top: 8px; }
            QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 4px; color: #111827; }
            QLineEdit, QSpinBox, QPlainTextEdit, QTextEdit, QTableWidget { background: #ffffff; color: #111827; border: 1px solid #c7d0db; border-radius: 6px; padding: 6px; selection-background-color: #bfdbfe; selection-color: #111827; }
            QLineEdit:focus, QSpinBox:focus, QPlainTextEdit:focus, QTextEdit:focus, QTableWidget:focus { border: 1px solid #2563eb; }
            QHeaderView::section { background: #f3f4f6; color: #111827; border: none; border-bottom: 1px solid #d1d5db; padding: 6px; font-weight: 600; }
            QTableWidget { gridline-color: #e5e7eb; alternate-background-color: #f9fafb; }
            QPushButton { background: #e5e7eb; color: #111827; border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px 12px; }
            QPushButton:hover:!disabled { background: #dbe4f0; }
            QPushButton:pressed:!disabled { background: #cfd8e3; }
            QPushButton:disabled { background: #f3f4f6; color: #9ca3af; border: 1px solid #e5e7eb; }
            QPushButton[role="primary"] { background: #2563eb; color: #ffffff; border: 1px solid #1d4ed8; font-weight: 600; }
            QPushButton[role="primary"]:hover:!disabled { background: #1d4ed8; }
            QPushButton[role="primary"]:pressed:!disabled { background: #1e40af; }
            QProgressBar { background: #e5e7eb; border: 1px solid #cbd5e1; border-radius: 7px; }
            QProgressBar::chunk { background: #2563eb; border-radius: 7px; }
            QLabel { color: #111827; background: transparent; }
            QLabel#stepStatus { color: #1d4ed8; font-weight: 600; }
            #StepCard { border: 1px solid #d0d7de; border-radius: 10px; background: #ffffff; }
            #StepCard:hover { border-color: #93c5fd; }
            #validationLabel { border-radius: 6px; }
            """
        )

    def switch_page(self, index: int) -> None:
        self.stack.setCurrentIndex(index)
        self.update_context_for_index(index)

    def focus_step(self, step_id: str) -> None:
        step_index = 1 + next(i for i, step in enumerate(STEPS) if step.step_id == step_id)
        self.nav_list.setCurrentRow(step_index)

    def update_context_for_index(self, index: int) -> None:
        if index == 0:
            self.context_step.setText("Overview")
            self.context_desc.setText("View workflow status and jump directly into any step.")
            self.context_status.setText("")
            self.context_output.setText("")
            self.context_log.clear()
            self.open_output_btn.setEnabled(False)
            return
        step_offset = index - 1
        if 0 <= step_offset < len(STEPS):
            step = STEPS[step_offset]
            step_state = self.state.step_data(step.step_id)
            self.context_step.setText(step.title)
            self.context_desc.setText(step.description)
            self.context_status.setText(f"Status: {step_state.get('status', 'not_configured')}")
            latest_output = self.find_latest_output(step.step_id)
            self.context_output.setText(f"Latest output: {latest_output or 'Not set'}")
            self.context_log.setPlainText(step_state.get("last_log", ""))
            self.open_output_btn.setEnabled(bool(latest_output))
            self.current_step_id = step.step_id
        else:
            labels = {len(STEPS) + 1: "Run History", len(STEPS) + 2: "Settings", len(STEPS) + 3: "Help / Run Guide"}
            self.context_step.setText(labels.get(index, APP_TITLE))
            self.context_desc.setText("")
            self.context_status.setText("")
            self.context_output.setText("")
            self.context_log.clear()
            self.open_output_btn.setEnabled(False)

    def find_latest_output(self, step_id: str) -> str:
        step_state = self.state.step_data(step_id)
        values = step_state.get("values", {})
        step = STEP_MAP[step_id]
        for output_key in step.outputs:
            if values.get(output_key):
                return values[output_key]
        return ""

    def run_step(self, step_id: str, payload: dict) -> None:
        if self.current_process is not None:
            QMessageBox.warning(self, "Workflow running", "Another step is already running.")
            self.statusBar().showMessage("A workflow step is already running.", 5000)
            return
        step = STEP_MAP[step_id]
        page = self.step_pages[step_id]
        page.log_edit.clear()
        page.set_running(True)
        page.summary_label.setText("Running...")
        page.save_to_state({"status": "running", "last_summary": "Running...", "values": payload["values"]})
        self.refresh_all()

        temp = tempfile.NamedTemporaryFile(delete=False, suffix=".json")
        temp_path = Path(temp.name)
        temp.close()
        temp_path.write_text(json.dumps({"responses": payload["responses"]}, indent=2), encoding="utf-8")
        self.current_temp_config = temp_path
        self.current_step_id = step_id
        self.current_log_buffer = ""

        process = QProcess(self)
        process.setWorkingDirectory(str(Path(__file__).resolve().parent))
        process.readyReadStandardOutput.connect(self.handle_stdout)
        process.readyReadStandardError.connect(self.handle_stderr)
        process.finished.connect(self.handle_finished)
        process.errorOccurred.connect(self.handle_process_error)
        env = QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONUNBUFFERED", "1")
        process.setProcessEnvironment(env)
        process.setProgram(sys.executable)
        process.setArguments(["-u", str(Path(__file__).with_name("ui_runner.py")), "--step", step_id, "--config", str(temp_path)])
        self.current_process = process
        process.start()
        self.statusBar().showMessage(f"Running {step.title}...")

    def handle_stdout(self) -> None:
        if not self.current_process or not self.current_step_id:
            return
        text = bytes(self.current_process.readAllStandardOutput()).decode(errors="replace")
        self.current_log_buffer += text
        self.step_pages[self.current_step_id].append_log(text)
        self.context_log.moveCursor(QTextCursor.End)
        self.context_log.insertPlainText(text)
        self.context_log.moveCursor(QTextCursor.End)
        self.state.step_data(self.current_step_id)["last_log"] = self.current_log_buffer

    def handle_stderr(self) -> None:
        if not self.current_process or not self.current_step_id:
            return
        text = bytes(self.current_process.readAllStandardError()).decode(errors="replace")
        self.current_log_buffer += text
        self.step_pages[self.current_step_id].append_log(text)
        self.context_log.moveCursor(QTextCursor.End)
        self.context_log.insertPlainText(text)
        self.context_log.moveCursor(QTextCursor.End)
        self.state.step_data(self.current_step_id)["last_log"] = self.current_log_buffer

    def handle_process_error(self, error) -> None:
        if not self.current_step_id:
            return
        self.step_pages[self.current_step_id].append_log(f"\nProcess error: {error}\n")

    def handle_finished(self, exit_code: int, exit_status) -> None:
        if not self.current_step_id:
            return
        step_id = self.current_step_id
        step = STEP_MAP[step_id]
        page = self.step_pages[step_id]
        log_text = self.current_log_buffer or page.log_edit.toPlainText()
        result = "success" if exit_code == 0 else "failed"
        latest_output = self.find_latest_output(step_id)
        summary = f"Run completed with exit code {exit_code}."
        if exit_code == 0 and latest_output:
            summary += f" Latest output: {latest_output}"
        page.summary_label.setText(summary)
        page.set_running(False)
        page.save_to_state(
            {
                "status": result,
                "last_run": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "last_log": log_text,
                "last_summary": summary,
            }
        )
        self.state.add_history(
            {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "step_id": step_id,
                "step_title": step.title,
                "result": result,
                "exit_code": exit_code,
                "latest_output": latest_output,
            }
        )
        if exit_code == 0 and self.state.data.get("settings", {}).get("auto_open_output_folder") and latest_output:
            target = Path(latest_output)
            folder = target if target.is_dir() else target.parent
            QDesktopServices.openUrl(folder.as_uri())

        if self.current_temp_config and self.current_temp_config.exists():
            try:
                self.current_temp_config.unlink()
            except Exception:
                pass
        self.statusBar().showMessage(f"{step.title} finished with result: {result}.", 7000)
        self.current_process = None
        self.current_step_id = None
        self.current_temp_config = None
        self.current_log_buffer = ""
        self.refresh_all()

    def refresh_all(self) -> None:
        for page in self.step_pages.values():
            page.refresh_prereq_status()
        self.overview_page.refresh()
        self.history_page.refresh()
        self.update_context_for_index(self.nav_list.currentRow())

    def open_latest_output(self) -> None:
        if not self.current_step_id:
            return
        latest_output = self.find_latest_output(self.current_step_id)
        if not latest_output:
            return
        target = Path(latest_output)
        folder = target if target.is_dir() else target.parent
        QDesktopServices.openUrl(folder.as_uri())


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
