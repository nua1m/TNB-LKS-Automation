import os
import re
import subprocess
import sys
import threading
from pathlib import Path

from config import DEFAULT_TEMPLATE_PATH
from main import run_process

try:
    from PySide6.QtCore import QObject, Qt, QThread, QUrl, Signal
    from PySide6.QtGui import QColor, QDesktopServices, QFont, QPalette
    from PySide6.QtWidgets import (
        QApplication,
        QFileDialog,
        QFrame,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPlainTextEdit,
        QPushButton,
        QSizePolicy,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:
    raise SystemExit(
        "PySide6 is required to run the desktop app. Install requirements.txt and launch again."
    ) from exc


APP_DIR = Path(__file__).resolve().parent
VERSION_FILE = APP_DIR / "VERSION"
VERSION = VERSION_FILE.read_text(encoding="utf-8").strip() if VERSION_FILE.exists() else "0.0.0"
ANSI_PATTERN = re.compile(r"\x1b\[[0-9;]*m")


class ProcessorWorker(QObject):
    log_message = Signal(str)
    status_changed = Signal(str)
    finished = Signal(dict)
    failed = Signal(str)
    append_confirmation_requested = Signal(int, int)

    def __init__(self, data_path: Path, template_path: Path):
        super().__init__()
        self.data_path = data_path
        self.template_path = template_path
        self._confirm_event = threading.Event()
        self._confirm_value = False

    def run(self) -> None:
        try:
            result = run_process(
                self.data_path,
                self.template_path,
                log_fn=self.log_message.emit,
                confirm_append_fn=self._confirm_append,
                status_fn=self.status_changed.emit,
                show_cli_summary=False,
            )
        except Exception as exc:
            self.failed.emit(str(exc))
            return

        self.finished.emit(result)

    def _confirm_append(self, existing_count: int, new_count: int) -> bool:
        self._confirm_event.clear()
        self.append_confirmation_requested.emit(existing_count, new_count)
        self._confirm_event.wait()
        return self._confirm_value

    def set_append_confirmation(self, answer: bool) -> None:
        self._confirm_value = answer
        self._confirm_event.set()


class StatChip(QFrame):
    def __init__(self, title: str, value: str):
        super().__init__()
        self.setObjectName("statChip")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(4)

        title_label = QLabel(title)
        title_label.setObjectName("statTitle")
        layout.addWidget(title_label)

        self.value_label = QLabel(value)
        self.value_label.setObjectName("statValue")
        layout.addWidget(self.value_label)

    def set_value(self, value: str) -> None:
        self.value_label.setText(value)


class LauncherWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.processing = False
        self.last_output_path: str | None = None
        self.generated_input_path: str | None = None
        self.result_folder = str(APP_DIR)
        self.worker_thread: QThread | None = None
        self.worker: ProcessorWorker | None = None

        self.setWindowTitle(f"TNB LKS Automation v{VERSION}")
        self.resize(1080, 760)
        self.setMinimumSize(960, 680)

        self._build_ui()
        self._apply_theme()
        self._set_status("Ready")

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(24, 24, 24, 24)
        root_layout.setSpacing(18)

        header_layout = QVBoxLayout()
        header_layout.setSpacing(4)

        title = QLabel("TNB LKS Automation")
        title.setObjectName("pageTitle")
        header_layout.addWidget(title)

        subtitle = QLabel(
            "Professional desktop workflow for preparing LKS workbooks while keeping the Excel COM process local."
        )
        subtitle.setObjectName("pageSubtitle")
        subtitle.setWordWrap(True)
        header_layout.addWidget(subtitle)

        root_layout.addLayout(header_layout)

        status_row = QHBoxLayout()
        status_row.setSpacing(12)

        self.status_chip = StatChip("Status", "Ready")
        self.version_chip = StatChip("Version", VERSION)
        self.update_chip = StatChip("Updates", "Check manually")

        status_row.addWidget(self.status_chip)
        status_row.addWidget(self.version_chip)
        status_row.addWidget(self.update_chip)
        root_layout.addLayout(status_row)

        content_row = QHBoxLayout()
        content_row.setSpacing(18)

        left_column = QVBoxLayout()
        left_column.setSpacing(18)
        content_row.addLayout(left_column, 3)

        left_column.addWidget(self._build_input_card())
        left_column.addWidget(self._build_actions_card())

        self.summary_card = self._build_summary_card()
        left_column.addWidget(self.summary_card)
        left_column.addStretch(1)

        self.log_card = self._build_log_card()
        content_row.addWidget(self.log_card, 4)

        root_layout.addLayout(content_row, 1)

    def _build_input_card(self) -> QGroupBox:
        card = QGroupBox("Input Excel File")
        layout = QVBoxLayout(card)
        layout.setSpacing(10)

        description = QLabel(
            "Choose the technician Excel file (.xls or .xlsx). The app will create the LKS result workbook in the same folder."
        )
        description.setObjectName("mutedText")
        description.setWordWrap(True)
        layout.addWidget(description)

        file_row = QHBoxLayout()
        file_row.setSpacing(10)

        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("Select an input workbook")
        self.file_edit.setClearButtonEnabled(True)
        file_row.addWidget(self.file_edit, 1)

        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.select_file)
        file_row.addWidget(self.browse_button)

        layout.addLayout(file_row)
        return card

    def _build_actions_card(self) -> QGroupBox:
        card = QGroupBox("Actions")
        layout = QVBoxLayout(card)
        layout.setSpacing(12)

        help_text = QLabel(
            "Run the LKS process, check for GitHub release updates, or open the generated files after the run finishes."
        )
        help_text.setObjectName("mutedText")
        help_text.setWordWrap(True)
        layout.addWidget(help_text)

        primary_row = QHBoxLayout()
        primary_row.setSpacing(10)

        self.process_button = QPushButton("Process LKS")
        self.process_button.setObjectName("primaryButton")
        self.process_button.clicked.connect(self.start_processing)
        primary_row.addWidget(self.process_button, 1)

        self.update_button = QPushButton("Check Updates")
        self.update_button.clicked.connect(self.check_for_updates)
        primary_row.addWidget(self.update_button)

        layout.addLayout(primary_row)

        secondary_row = QHBoxLayout()
        secondary_row.setSpacing(10)

        self.open_folder_button = QPushButton("Open Result Folder")
        self.open_folder_button.clicked.connect(self.open_result_folder)
        secondary_row.addWidget(self.open_folder_button)

        self.open_result_file_button = QPushButton("Open Result File")
        self.open_result_file_button.clicked.connect(self.open_result_file)
        self.open_result_file_button.setEnabled(False)
        secondary_row.addWidget(self.open_result_file_button)

        self.open_generated_input_button = QPushButton("Open Cleaned Input")
        self.open_generated_input_button.clicked.connect(self.open_generated_input_file)
        self.open_generated_input_button.setEnabled(False)
        secondary_row.addWidget(self.open_generated_input_button)

        layout.addLayout(secondary_row)
        return card

    def _build_summary_card(self) -> QGroupBox:
        card = QGroupBox("Run Summary")
        layout = QVBoxLayout(card)
        layout.setSpacing(8)

        self.summary_label = QLabel(
            "No run completed yet. After processing, this panel will show the latest totals."
        )
        self.summary_label.setObjectName("summaryText")
        self.summary_label.setWordWrap(True)
        self.summary_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.summary_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(self.summary_label)

        return card

    def _build_log_card(self) -> QGroupBox:
        card = QGroupBox("Run Log")
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout = QVBoxLayout(card)
        layout.setSpacing(10)

        description = QLabel(
            "Live processing messages, warnings, and review notes appear here."
        )
        description.setObjectName("mutedText")
        description.setWordWrap(True)
        layout.addWidget(description)

        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setObjectName("logPane")
        self.log_text.setFont(QFont("Consolas", 10))
        layout.addWidget(self.log_text, 1)

        return card

    def _apply_theme(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                background: #f3f6fb;
                color: #132238;
                font-family: "Segoe UI";
                font-size: 10pt;
            }
            QMainWindow {
                background: #f3f6fb;
            }
            QLabel#pageTitle {
                font-size: 22pt;
                font-weight: 700;
                color: #0f172a;
            }
            QLabel#pageSubtitle {
                color: #536377;
                font-size: 10.5pt;
            }
            QGroupBox {
                background: white;
                border: 1px solid #d7e0ea;
                border-radius: 14px;
                font-weight: 600;
                margin-top: 12px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 16px;
                padding: 0 4px;
                color: #10233c;
            }
            QFrame#statChip {
                background: white;
                border: 1px solid #d7e0ea;
                border-radius: 14px;
            }
            QLabel#statTitle {
                color: #607086;
                font-size: 9pt;
            }
            QLabel#statValue {
                color: #0f172a;
                font-size: 16pt;
                font-weight: 700;
            }
            QLabel#mutedText {
                color: #607086;
            }
            QLabel#summaryText {
                color: #223247;
                background: #f8fbff;
                border: 1px solid #dbe6f1;
                border-radius: 10px;
                padding: 12px;
            }
            QLineEdit, QPlainTextEdit {
                background: #fbfdff;
                border: 1px solid #cfdae6;
                border-radius: 10px;
                padding: 10px 12px;
            }
            QLineEdit:focus, QPlainTextEdit:focus {
                border: 1px solid #2c6bed;
            }
            QPushButton {
                background: #eef3f9;
                border: 1px solid #cfdae6;
                border-radius: 10px;
                padding: 10px 14px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #e5edf7;
            }
            QPushButton:disabled {
                background: #f5f7fa;
                color: #9aa7b6;
                border-color: #e0e6ed;
            }
            QPushButton#primaryButton {
                background: #18794e;
                border: 1px solid #18794e;
                color: white;
            }
            QPushButton#primaryButton:hover {
                background: #146a44;
            }
            QPlainTextEdit#logPane {
                background: #0f172a;
                color: #d8e6ff;
                border: 1px solid #0f172a;
            }
            """
        )

    def _set_status(self, text: str) -> None:
        self.status_chip.set_value(text)

    def _set_update_text(self, text: str) -> None:
        self.update_chip.set_value(text)

    def append_log(self, message: str) -> None:
        clean_message = ANSI_PATTERN.sub("", message).rstrip()
        self.log_text.appendPlainText(clean_message)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    @staticmethod
    def format_error_message(raw_message: str) -> str:
        text = raw_message.strip()
        lowered = text.lower()

        if "missing '3ms so no.'" in lowered:
            return "The required column '3MS SO No.' was not found in the input file."
        if "template file not found" in lowered:
            return "The template file could not be found."
        if "permission" in lowered:
            return "The app could not access one of the files. Close Excel or any open workbook and try again."
        if "excel" in lowered and "could not auto-refresh" in lowered:
            return "The file was saved, but Excel could not refresh it automatically."
        return text

    def select_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Input Excel File",
            str(Path.home()),
            "Excel files (*.xls *.xlsx);;All files (*.*)",
        )
        if file_path:
            self.file_edit.setText(file_path)

    def set_processing_state(self, active: bool) -> None:
        self.processing = active
        self.process_button.setEnabled(not active)
        self.browse_button.setEnabled(not active)
        self.update_button.setEnabled(not active)
        self.file_edit.setEnabled(not active)

        if active:
            self.open_result_file_button.setEnabled(False)
            self.open_generated_input_button.setEnabled(False)

    def check_for_updates(self) -> None:
        updater_script = APP_DIR / "updater.py"
        try:
            subprocess.Popen(
                [sys.executable, str(updater_script), "--check-only"],
                cwd=str(APP_DIR),
            )
            self._set_update_text("Manual check started")
        except Exception as exc:
            QMessageBox.critical(self, "Update Check Failed", str(exc))

    def _open_path(self, target: str, title: str, missing_message: str) -> None:
        if not target or not os.path.exists(target):
            QMessageBox.warning(self, title, missing_message)
            return

        if not QDesktopServices.openUrl(QUrl.fromLocalFile(target)):
            QMessageBox.critical(self, title, f"Could not open:\n{target}")

    def open_result_folder(self) -> None:
        target = self.result_folder if os.path.isdir(self.result_folder) else str(APP_DIR)
        self._open_path(target, "Open Folder Failed", "The result folder could not be found.")

    def open_result_file(self) -> None:
        self._open_path(
            self.last_output_path or "",
            "Open Result File Failed",
            "The saved workbook could not be found.",
        )

    def open_generated_input_file(self) -> None:
        self._open_path(
            self.generated_input_path or "",
            "Open Input File Failed",
            "The generated input file could not be found.",
        )

    def start_processing(self) -> None:
        data_path = self.file_edit.text().strip()
        if not data_path:
            QMessageBox.warning(self, "Input File Required", "Choose the Excel file you want to process first.")
            return

        if not os.path.exists(data_path):
            QMessageBox.critical(self, "File Not Found", f"Could not find:\n{data_path}")
            return

        template_path = Path(DEFAULT_TEMPLATE_PATH).resolve()
        if not template_path.exists():
            QMessageBox.critical(self, "Template Not Found", f"Could not find the template file:\n{template_path}")
            return

        self.last_output_path = None
        self.generated_input_path = None
        self.result_folder = str(Path(data_path).resolve().parent)
        self.summary_label.setText("Run in progress. Summary will appear here after processing completes.")
        self.log_text.clear()
        self.append_log("Starting LKS processing...")
        self._set_status("Processing workbook...")
        self.set_processing_state(True)

        self.worker_thread = QThread(self)
        self.worker = ProcessorWorker(Path(data_path).resolve(), template_path)
        self.worker.moveToThread(self.worker_thread)

        self.worker.log_message.connect(self.append_log)
        self.worker.status_changed.connect(self._set_status)
        self.worker.append_confirmation_requested.connect(self._handle_append_confirmation)
        self.worker.finished.connect(self._handle_done)
        self.worker.failed.connect(self._handle_error)

        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.started.connect(self.worker.run)
        self.worker_thread.start()

    def _handle_append_confirmation(self, existing_count: int, new_count: int) -> None:
        if not self.worker:
            return

        answer = QMessageBox.question(
            self,
            "Append New SOs",
            (
                f"The template already has {existing_count} SOs.\n\n"
                f"{new_count} new SOs will be added.\n\n"
                "Choose Yes to continue. No changes will be saved if you choose No."
            ),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        self.worker.set_append_confirmation(answer == QMessageBox.Yes)

    def _handle_done(self, result: dict) -> None:
        self.set_processing_state(False)
        self.worker = None
        self.worker_thread = None

        self.last_output_path = result.get("output_path")
        self.generated_input_path = result.get("generated_input_path")
        if self.last_output_path:
            self.result_folder = str(Path(self.last_output_path).parent)

        if result.get("aborted"):
            self._set_status("Processing cancelled")
            self.append_log("Run cancelled before saving changes.")
            self.summary_label.setText("Run cancelled before saving changes.")
            return

        self._set_status("Completed")
        if self.last_output_path and os.path.exists(self.last_output_path):
            self.open_result_file_button.setEnabled(True)
        if self.generated_input_path and os.path.exists(self.generated_input_path):
            self.open_generated_input_button.setEnabled(True)

        if self.last_output_path:
            self.append_log(f"Saved to: {self.last_output_path}")
        if self.generated_input_path and self.generated_input_path != self.file_edit.text().strip():
            self.append_log(f"Cleaned input: {self.generated_input_path}")

        self._render_summary(result)

        QMessageBox.information(
            self,
            "Processing Complete",
            (
                "LKS processing completed successfully.\n\n"
                f"Saved to:\n{self.last_output_path}\n\n"
                "You can now open the result file, open the cleaned input file, or review the run log."
            ),
        )

    def _handle_error(self, error_message: str) -> None:
        self.set_processing_state(False)
        self.worker = None
        self.worker_thread = None

        self._set_status("Failed")
        friendly = self.format_error_message(error_message)
        self.append_log(f"Error: {friendly}")
        self.summary_label.setText("Run failed. Review the log and fix the issue before running again.")

        QMessageBox.critical(
            self,
            "Processing Failed",
            (
                "The file could not be processed.\n\n"
                f"{friendly}\n\n"
                "Check the run log for details."
            ),
        )

    def _render_summary(self, result: dict) -> None:
        summary = result.get("summary", {})
        tras_by_date = result.get("tras_by_date", {})

        if summary:
            self.append_log("")
            self.append_log("Summary:")
            for key, value in summary.items():
                self.append_log(f"- {key}: {value}")

        if tras_by_date:
            self.append_log("")
            self.append_log("TRAS by date:")
            for tras_date, tras_count in tras_by_date.items():
                self.append_log(f"- {tras_date}: {tras_count}")

        lines = []
        for key, value in summary.items():
            lines.append(f"{key}: {value}")

        if tras_by_date:
            lines.append("")
            lines.append("TRAS by date:")
            for tras_date, tras_count in tras_by_date.items():
                lines.append(f"- {tras_date}: {tras_count}")

        if not lines:
            lines.append("Run completed, but no summary details were returned.")

        self.summary_label.setText("\n".join(lines))


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("TNB LKS Automation")
    app.setOrganizationName("TNB LKS")
    app.setStyle("Fusion")

    palette = app.palette()
    palette.setColor(QPalette.Window, QColor("#f3f6fb"))
    palette.setColor(QPalette.Base, QColor("#fbfdff"))
    palette.setColor(QPalette.Button, QColor("#eef3f9"))
    app.setPalette(palette)

    window = LauncherWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
