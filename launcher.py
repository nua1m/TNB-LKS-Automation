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
        QDialog,
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


class InfoDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("App Information")
        self.setModal(True)
        self.setMinimumWidth(320)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        title = QLabel("App Information")
        title.setObjectName("infoDialogTitle")
        layout.addWidget(title)

        self.version_value = QLabel("")
        self.version_value.setObjectName("infoValue")
        self.email_value = QLabel("syahmi@nuaim.my")
        self.email_value.setObjectName("infoValue")
        self.phone_value = QLabel("+60 18 2605 390")
        self.phone_value.setObjectName("infoValue")

        for label_text, value_widget in (
            ("Version", self.version_value),
            ("Support Email", self.email_value),
            ("Support Phone", self.phone_value),
        ):
            row = QVBoxLayout()
            row.setSpacing(2)
            label = QLabel(label_text)
            label.setObjectName("infoLabel")
            row.addWidget(label)
            row.addWidget(value_widget)
            layout.addLayout(row)

        self.check_updates_button = QPushButton("Check Updates")
        layout.addWidget(self.check_updates_button)

    def set_values(self, version: str) -> None:
        self.version_value.setText(version)


class FileDropArea(QFrame):
    file_dropped = Signal(str)

    def __init__(self):
        super().__init__()
        self.setObjectName("fileDropArea")
        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(8)
        layout.setAlignment(Qt.AlignCenter)

        self.icon_label = QLabel("Upload")
        self.icon_label.setObjectName("dropTitle")
        self.icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.icon_label)

        self.text_label = QLabel("Drag and drop your Excel file here")
        self.text_label.setObjectName("dropText")
        self.text_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.text_label)

        self.hint_label = QLabel("Supports .xls and .xlsx")
        self.hint_label.setObjectName("dropHint")
        self.hint_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.hint_label)

        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.MinimumExpanding)

    @staticmethod
    def _extract_supported_file(event) -> str | None:
        mime_data = event.mimeData()
        if not mime_data.hasUrls():
            return None

        for url in mime_data.urls():
            if not url.isLocalFile():
                continue
            local_path = url.toLocalFile()
            if local_path.lower().endswith((".xls", ".xlsx")):
                return local_path
        return None

    def dragEnterEvent(self, event) -> None:
        file_path = self._extract_supported_file(event)
        if file_path:
            event.acceptProposedAction()
            self.setProperty("dragActive", True)
            self.style().unpolish(self)
            self.style().polish(self)
            self.update()
            return
        event.ignore()

    def dragLeaveEvent(self, event) -> None:
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        self.update()
        super().dragLeaveEvent(event)

    def dropEvent(self, event) -> None:
        file_path = self._extract_supported_file(event)
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        self.update()

        if file_path:
            self.file_dropped.emit(file_path)
            event.acceptProposedAction()
            return
        event.ignore()


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
        self.resize(1180, 780)
        self.setMinimumSize(1020, 700)

        self._build_ui()
        self._apply_theme()
        self.info_dialog = InfoDialog(self)
        self.info_dialog.check_updates_button.clicked.connect(self.check_for_updates)
        self._set_status("Ready")
        self._refresh_info_dialog()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(24, 24, 24, 24)
        root_layout.setSpacing(18)

        header_layout = QVBoxLayout()
        header_layout.setSpacing(8)

        top_row = QHBoxLayout()
        top_row.setSpacing(12)

        title = QLabel("TNB LKS Automation")
        title.setObjectName("pageTitle")
        top_row.addWidget(title)
        top_row.addStretch(1)

        self.info_button = QPushButton("i")
        self.info_button.setObjectName("infoButton")
        self.info_button.setToolTip("View app information")
        self.info_button.clicked.connect(self.show_info_dialog)
        top_row.addWidget(self.info_button)
        header_layout.addLayout(top_row)

        root_layout.addLayout(header_layout)

        left_panel = QWidget()
        left_column = QVBoxLayout(left_panel)
        left_column.setContentsMargins(0, 0, 0, 0)
        left_column.setSpacing(18)
        left_column.addWidget(self._build_input_card(), 1)
        left_column.addWidget(self._build_actions_card())
        self.summary_card = self._build_summary_card()
        left_column.addWidget(self.summary_card, 1)

        right_panel = QWidget()
        right_column = QVBoxLayout(right_panel)
        right_column.setContentsMargins(0, 0, 0, 0)
        right_column.setSpacing(18)

        self.log_card = self._build_log_card()

        right_column.addWidget(self.log_card, 1)

        content_row = QHBoxLayout()
        content_row.setSpacing(18)
        content_row.addWidget(left_panel, 3)
        content_row.addWidget(right_panel, 4)
        root_layout.addLayout(content_row, 1)

    def _build_input_card(self) -> QGroupBox:
        card = QGroupBox("Input Excel File")
        layout = QVBoxLayout(card)
        layout.setSpacing(12)

        self.drop_area = FileDropArea()
        self.drop_area.setMinimumHeight(140)
        self.drop_area.setMaximumHeight(240)
        self.drop_area.file_dropped.connect(self._set_selected_file)
        layout.addWidget(self.drop_area, 1)

        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("Select an input workbook")
        self.file_edit.setClearButtonEnabled(True)
        layout.addWidget(self.file_edit)

        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.select_file)
        layout.addWidget(self.browse_button)
        return card

    def _build_actions_card(self) -> QGroupBox:
        card = QGroupBox("Actions")
        layout = QVBoxLayout(card)
        layout.setSpacing(12)

        primary_row = QHBoxLayout()
        primary_row.setSpacing(10)

        self.process_button = QPushButton("Process LKS")
        self.process_button.setObjectName("primaryButton")
        self.process_button.clicked.connect(self.start_processing)
        primary_row.addWidget(self.process_button, 1)

        layout.addLayout(primary_row)

        secondary_row = QHBoxLayout()
        secondary_row.setSpacing(10)

        self.open_folder_button = QPushButton("Open Result Folder")
        self.open_folder_button.clicked.connect(self.open_result_folder)
        secondary_row.addWidget(self.open_folder_button)

        self.open_result_file_button = QPushButton("Open LKS")
        self.open_result_file_button.clicked.connect(self.open_result_file)
        self.open_result_file_button.setEnabled(False)
        secondary_row.addWidget(self.open_result_file_button)

        self.open_generated_input_button = QPushButton("Open Raw Data")
        self.open_generated_input_button.clicked.connect(self.open_generated_input_file)
        self.open_generated_input_button.setEnabled(False)
        secondary_row.addWidget(self.open_generated_input_button)

        layout.addLayout(secondary_row)
        return card

    def _build_summary_card(self) -> QGroupBox:
        card = QGroupBox("Run Summary")
        layout = QVBoxLayout(card)
        layout.setSpacing(8)

        self.summary_text = QPlainTextEdit()
        self.summary_text.setObjectName("summaryPane")
        self.summary_text.setReadOnly(True)
        self.summary_text.setPlainText(
            "No run completed yet. After processing, this panel will show the latest totals."
        )
        self.summary_text.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(self.summary_text, 1)

        return card

    def _build_log_card(self) -> QGroupBox:
        card = QGroupBox("Activity")
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout = QVBoxLayout(card)
        layout.setSpacing(10)

        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setObjectName("logPane")
        self.log_text.setFont(QFont("Inter", 10))
        layout.addWidget(self.log_text, 1)

        return card

    def _apply_theme(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                background: #f6f7f9;
                color: #182230;
                font-family: "Inter", "Segoe UI";
                font-size: 10pt;
            }
            QMainWindow {
                background: #f6f7f9;
            }
            QLabel#pageTitle {
                font-size: 22pt;
                font-weight: 700;
                color: #111827;
            }
            QFrame#fileDropArea {
                background: #fbfcfd;
                border: 2px dashed #d7e0ea;
                border-radius: 12px;
            }
            QFrame#fileDropArea[dragActive="true"] {
                background: #f2f7ff;
                border: 2px dashed #7aa2e3;
            }
            QLabel#dropTitle {
                color: #111827;
                font-size: 12pt;
                font-weight: 700;
            }
            QLabel#dropText {
                color: #334155;
                font-size: 10pt;
                font-weight: 600;
            }
            QLabel#dropHint {
                color: #6b7280;
                font-size: 9pt;
            }
            QGroupBox {
                background: white;
                border: 1px solid #e4e9ef;
                border-radius: 10px;
                font-weight: 600;
                margin-top: 10px;
                padding-top: 14px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 16px;
                padding: 0 4px;
                color: #1f2937;
            }
            QLabel#mutedText {
                color: #667085;
            }
            QPushButton#infoButton {
                min-width: 32px;
                max-width: 32px;
                min-height: 32px;
                max-height: 32px;
                background: #ffffff;
                border: 1px solid #dfe5ec;
                border-radius: 10px;
                color: #4b5563;
                font-size: 12pt;
                font-weight: 700;
                padding: 0px;
            }
            QPushButton#infoButton:hover {
                background: #f6f8fb;
            }
            QLineEdit, QPlainTextEdit {
                background: #fcfdff;
                border: 1px solid #dde4ec;
                border-radius: 8px;
                padding: 11px 13px;
            }
            QLineEdit:focus, QPlainTextEdit:focus {
                border: 1px solid #8aa6d6;
            }
            QPushButton {
                background: #ffffff;
                border: 1px solid #dde4ec;
                border-radius: 8px;
                padding: 10px 15px;
                font-weight: 600;
                color: #243041;
            }
            QPushButton:hover {
                background: #f6f8fb;
            }
            QPushButton:disabled {
                background: #f8fafc;
                color: #9aa5b1;
                border-color: #e5eaf0;
            }
            QPushButton#primaryButton {
                background: #1f7a4f;
                border: 1px solid #1f7a4f;
                color: white;
            }
            QPushButton#primaryButton:hover {
                background: #1a6844;
            }
            QPlainTextEdit#summaryPane {
                background: #fcfdff;
                color: #243041;
                border: 1px solid #e1e7ee;
                font-size: 10pt;
            }
            QPlainTextEdit#logPane {
                background: #fcfdff;
                color: #344054;
                border: 1px solid #e1e7ee;
            }
            QScrollBar:vertical {
                background: transparent;
                width: 12px;
                margin: 8px 3px 8px 3px;
            }
            QScrollBar::handle:vertical {
                background: #c4cfdb;
                min-height: 36px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical:hover {
                background: #aab9ca;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: transparent;
            }
            QLabel#infoDialogTitle {
                color: #1f2937;
                font-size: 12pt;
                font-weight: 700;
            }
            QLabel#infoLabel {
                color: #667085;
                font-size: 9pt;
                font-weight: 600;
            }
            QLabel#infoValue {
                color: #1f2937;
                font-size: 10pt;
            }
            """
        )

    def _set_status(self, text: str) -> None:
        return None

    def _set_update_text(self, text: str) -> None:
        return None

    def _refresh_info_dialog(self) -> None:
        self.info_dialog.set_values(version=VERSION)

    def show_info_dialog(self) -> None:
        self._refresh_info_dialog()
        self.info_dialog.exec()

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
            self._set_selected_file(file_path)

    def _set_selected_file(self, file_path: str) -> None:
        self.file_edit.setText(file_path)
        file_name = Path(file_path).name
        self.drop_area.text_label.setText(file_name)
        self.drop_area.hint_label.setText(str(Path(file_path)))

    def set_processing_state(self, active: bool) -> None:
        self.processing = active
        self.process_button.setEnabled(not active)
        self.browse_button.setEnabled(not active)
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
            "Open LKS Failed",
            "The saved LKS workbook could not be found.",
        )

    def open_generated_input_file(self) -> None:
        self._open_path(
            self.generated_input_path or "",
            "Open Raw Data Failed",
            "The processed raw data file could not be found.",
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
        self.summary_text.setPlainText("Run in progress. Summary will appear here after processing completes.")
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
            self.summary_text.setPlainText("Run cancelled before saving changes.")
            return

        self._set_status("Completed")
        if self.last_output_path and os.path.exists(self.last_output_path):
            self.open_result_file_button.setEnabled(True)
        if self.generated_input_path and os.path.exists(self.generated_input_path):
            self.open_generated_input_button.setEnabled(True)

        if self.last_output_path:
            self.append_log(f"Saved to: {self.last_output_path}")
        if self.generated_input_path and self.generated_input_path != self.file_edit.text().strip():
            self.append_log(f"Raw data file: {self.generated_input_path}")

        self._render_summary(result)

        QMessageBox.information(
            self,
            "Processing Complete",
            (
                "LKS processing completed successfully.\n\n"
                f"Saved to:\n{self.last_output_path}\n\n"
                "You can now open the LKS file, open the raw data file, or review the run log."
            ),
        )

    def _handle_error(self, error_message: str) -> None:
        self.set_processing_state(False)
        self.worker = None
        self.worker_thread = None

        self._set_status("Failed")
        friendly = self.format_error_message(error_message)
        self.append_log(f"Error: {friendly}")
        self.summary_text.setPlainText("Run failed. Review the log and fix the issue before running again.")

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

        self.summary_text.setPlainText("\n".join(lines))


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
    window.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
