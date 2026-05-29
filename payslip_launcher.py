from __future__ import annotations

import subprocess
import sys
import traceback
from datetime import date
from pathlib import Path

from updater import APP_DIR
from ui_theme import apply_app_palette, panel_stylesheet

try:
    from PySide6.QtCore import QObject, QDate, QThread, QUrl, Signal
    from PySide6.QtGui import QDesktopServices, QFont
    from PySide6.QtWidgets import (
        QApplication,
        QDateEdit,
        QDialog,
        QFileDialog,
        QFormLayout,
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


VERSION_FILE = APP_DIR / "VERSION"
VERSION = VERSION_FILE.read_text(encoding="utf-8").strip() if VERSION_FILE.exists() else "0.0.0"
ERROR_LOG_FILE = APP_DIR / "payslip_launcher_error.log"


def _report_startup_error(title: str, exc: Exception) -> None:
    details = traceback.format_exc()
    ERROR_LOG_FILE.write_text(details, encoding="utf-8")

    try:
        from tkinter import Tk, messagebox

        root = Tk()
        root.withdraw()
        messagebox.showerror(
            title,
            f"{exc}\n\nA full error log was saved to:\n{ERROR_LOG_FILE}",
        )
        root.destroy()
    except Exception:
        pass

    raise SystemExit(1) from exc


try:
    from core.services.payslip_service import (
        DEFAULT_CALC_PATH,
        DEFAULT_LKS_SAMPLE_PATH,
        DEFAULT_MASTER_PATH,
        DEFAULT_OUTPUT_DIR,
        generate_payslips,
    )
except Exception as exc:  # pragma: no cover - startup diagnostics
    _report_startup_error("Payslip Generator Failed To Start", exc)


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


class PayslipWorker(QObject):
    log_message = Signal(str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        calc_path: Path,
        master_path: Path,
        output_dir: Path,
        salary_month: str,
        payment_date: date,
        lks_paths: list[Path],
    ):
        super().__init__()
        self.calc_path = calc_path
        self.master_path = master_path
        self.output_dir = output_dir
        self.salary_month = salary_month
        self.payment_date = payment_date
        self.lks_paths = lks_paths

    def run(self) -> None:
        try:
            if self.lks_paths:
                self.log_message.emit("Loading LKS CLAIM workbooks...")
                self.log_message.emit("Building calculation workbook from CLAIM rows...")
            else:
                self.log_message.emit("Loading calculation workbook...")
            self.log_message.emit("Loading worker master file...")
            result = generate_payslips(
                calc_path=self.calc_path,
                master_path=self.master_path,
                output_dir=self.output_dir,
                salary_month=self.salary_month,
                payment_date=self.payment_date,
                lks_paths=self.lks_paths,
            )
        except Exception as exc:
            self.failed.emit(str(exc))
            return

        self.finished.emit(result)


class PayslipPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.processing = False
        self.worker_thread: QThread | None = None
        self.worker: PayslipWorker | None = None
        self.last_output_dir: Path | None = None
        self.selected_lks_paths: list[Path] = []

        self._build_ui()
        self._apply_theme()

        self.info_dialog = InfoDialog(self)
        self.info_dialog.set_values(VERSION)
        self.info_dialog.check_updates_button.clicked.connect(self.check_for_updates)

        if DEFAULT_CALC_PATH.exists():
            self.calc_edit.setText(str(DEFAULT_CALC_PATH))
        if DEFAULT_MASTER_PATH.exists():
            self.master_edit.setText(str(DEFAULT_MASTER_PATH))
        self.output_edit.setText(str(DEFAULT_OUTPUT_DIR))
        self.month_edit.setText(QDate.currentDate().toString("MMMM yyyy"))
        self.payment_date_edit.setDate(QDate.currentDate())

    def _build_ui(self) -> None:
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(24, 24, 24, 24)
        root_layout.setSpacing(18)

        top_row = QHBoxLayout()
        title = QLabel("TNB Payslip Generator")
        title.setObjectName("pageTitle")
        top_row.addWidget(title)
        top_row.addStretch(1)

        self.info_button = QPushButton("i")
        self.info_button.setObjectName("infoButton")
        self.info_button.clicked.connect(self.show_info_dialog)
        top_row.addWidget(self.info_button)
        root_layout.addLayout(top_row)

        content_row = QHBoxLayout()
        content_row.setSpacing(18)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(18)
        left_layout.addWidget(self._build_inputs_card())
        left_layout.addWidget(self._build_actions_card())
        left_layout.addWidget(self._build_summary_card(), 1)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(18)
        right_layout.addWidget(self._build_log_card(), 1)

        content_row.addWidget(left_panel, 3)
        content_row.addWidget(right_panel, 4)
        root_layout.addLayout(content_row, 1)

    def _build_inputs_card(self) -> QGroupBox:
        card = QGroupBox("Generation Inputs")
        layout = QFormLayout(card)
        layout.setSpacing(12)
        layout.setLabelAlignment(layout.labelAlignment())

        self.calc_edit, calc_row, calc_button = self._path_row("Select the filled TNB calculation workbook")
        self.master_edit, master_row, master_button = self._path_row("Select the worker/team master workbook")
        self.lks_edit, lks_row, lks_button = self._path_row("Optional: select one or more LKS files")
        self.output_edit, output_row, output_button = self._path_row("Select where generated payslips should be saved")
        self.lks_edit.setReadOnly(True)

        calc_button.clicked.connect(self.select_calc_file)
        master_button.clicked.connect(self.select_master_file)
        lks_button.clicked.connect(self.select_lks_files)
        output_button.clicked.connect(self.select_output_dir)

        self.month_edit = QLineEdit()
        self.month_edit.setPlaceholderText("Example: May 2026")

        self.payment_date_edit = QDateEdit()
        self.payment_date_edit.setCalendarPopup(True)
        self.payment_date_edit.setDisplayFormat("dd MMM yyyy")

        layout.addRow("Calculation Workbook", calc_row)
        layout.addRow("Worker Master File", master_row)
        layout.addRow("LKS Files", lks_row)
        layout.addRow("Salary Month", self.month_edit)
        layout.addRow("Payment Date", self.payment_date_edit)
        layout.addRow("Output Folder", output_row)
        return card

    def _build_actions_card(self) -> QGroupBox:
        card = QGroupBox("Actions")
        layout = QVBoxLayout(card)
        layout.setSpacing(12)

        primary_row = QHBoxLayout()
        primary_row.setSpacing(10)

        self.generate_button = QPushButton("Generate Payslips")
        self.generate_button.setObjectName("primaryButton")
        self.generate_button.clicked.connect(self.start_generation)
        primary_row.addWidget(self.generate_button, 1)

        layout.addLayout(primary_row)

        secondary_row = QHBoxLayout()
        secondary_row.setSpacing(10)

        self.clear_lks_button = QPushButton("Clear LKS Files")
        self.clear_lks_button.setEnabled(False)
        self.clear_lks_button.clicked.connect(self.clear_lks_files)
        secondary_row.addWidget(self.clear_lks_button)

        self.open_output_button = QPushButton("Open Output Folder")
        self.open_output_button.setEnabled(False)
        self.open_output_button.clicked.connect(self.open_output_folder)
        secondary_row.addWidget(self.open_output_button)

        layout.addLayout(secondary_row)
        return card

    def _build_summary_card(self) -> QGroupBox:
        card = QGroupBox("Run Summary")
        layout = QVBoxLayout(card)
        self.summary_text = QPlainTextEdit()
        self.summary_text.setObjectName("summaryPane")
        self.summary_text.setReadOnly(True)
        self.summary_text.setPlainText(
            "No run completed yet. After generation, this panel will show generated files and warnings."
        )
        layout.addWidget(self.summary_text, 1)
        return card

    def _build_log_card(self) -> QGroupBox:
        card = QGroupBox("Activity")
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout = QVBoxLayout(card)
        self.log_text = QPlainTextEdit()
        self.log_text.setObjectName("logPane")
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Inter", 10))
        layout.addWidget(self.log_text, 1)
        return card

    def _path_row(self, placeholder: str) -> tuple[QLineEdit, QWidget, QPushButton]:
        container = QWidget()
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)

        line_edit = QLineEdit()
        line_edit.setPlaceholderText(placeholder)
        row.addWidget(line_edit, 1)

        button = QPushButton("Browse")
        row.addWidget(button)
        return line_edit, container, button

    def _refresh_lks_display(self) -> None:
        if not self.selected_lks_paths:
            self.lks_edit.clear()
            self.lks_edit.setToolTip("")
            self.clear_lks_button.setEnabled(False)
            return

        if len(self.selected_lks_paths) == 1:
            self.lks_edit.setText(str(self.selected_lks_paths[0]))
        else:
            self.lks_edit.setText(f"{len(self.selected_lks_paths)} files selected")
        self.lks_edit.setToolTip("\n".join(str(path) for path in self.selected_lks_paths))
        self.clear_lks_button.setEnabled(True)

    def _apply_theme(self) -> None:
        self.setStyleSheet(panel_stylesheet())

    def show_info_dialog(self) -> None:
        self.info_dialog.exec()

    def append_log(self, message: str) -> None:
        self.log_text.appendPlainText(message.rstrip())
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def set_processing_state(self, active: bool) -> None:
        self.processing = active
        self.generate_button.setEnabled(not active)
        self.open_output_button.setEnabled((not active) and self.last_output_dir is not None)
        for widget in (
            self.calc_edit,
            self.master_edit,
            self.lks_edit,
            self.output_edit,
            self.month_edit,
            self.payment_date_edit,
        ):
            widget.setEnabled(not active)

    def select_calc_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Calculation Workbook",
            str(DEFAULT_CALC_PATH.parent),
            "Excel files (*.xlsx *.xls);;All files (*.*)",
        )
        if file_path:
            self.calc_edit.setText(file_path)

    def select_master_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Worker Master Workbook",
            str(DEFAULT_MASTER_PATH.parent),
            "Excel files (*.xlsx *.xls);;All files (*.*)",
        )
        if file_path:
            self.master_edit.setText(file_path)

    def select_output_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder",
            self.output_edit.text() or str(DEFAULT_OUTPUT_DIR),
        )
        if folder:
            self.output_edit.setText(folder)

    def select_lks_files(self) -> None:
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select LKS Files",
            str(DEFAULT_LKS_SAMPLE_PATH.parent),
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        if file_paths:
            existing = {str(path).lower(): path for path in self.selected_lks_paths}
            for file_path in file_paths:
                path = Path(file_path)
                existing[str(path).lower()] = path
            self.selected_lks_paths = list(existing.values())
            self._refresh_lks_display()

    def clear_lks_files(self) -> None:
        self.selected_lks_paths = []
        self._refresh_lks_display()

    def check_for_updates(self) -> None:
        updater_script = APP_DIR / "updater.py"
        subprocess.Popen([sys.executable, str(updater_script), "--check-only"], cwd=str(APP_DIR))

    def open_output_folder(self) -> None:
        if self.last_output_dir is None:
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_output_dir)))

    def start_generation(self) -> None:
        calc_path = Path(self.calc_edit.text().strip())
        master_path = Path(self.master_edit.text().strip())
        output_dir = Path(self.output_edit.text().strip()) if self.output_edit.text().strip() else DEFAULT_OUTPUT_DIR
        salary_month = self.month_edit.text().strip()
        payment_date = self.payment_date_edit.date().toPython()
        lks_paths = list(self.selected_lks_paths)

        if not calc_path.exists():
            QMessageBox.warning(self, "Missing File", "Select a valid calculation workbook.")
            return
        if not master_path.exists():
            QMessageBox.warning(self, "Missing File", "Select a valid worker master workbook.")
            return
        for path in lks_paths:
            if not path.exists():
                QMessageBox.warning(self, "Missing File", f"LKS file not found:\n{path}")
                return
        if not salary_month:
            QMessageBox.warning(self, "Missing Month", "Enter the salary month label.")
            return

        self.log_text.clear()
        self.summary_text.setPlainText("Generating payslips...")
        self.last_output_dir = None
        self.set_processing_state(True)
        self.append_log("Starting payslip generation...")

        self.worker_thread = QThread(self)
        self.worker = PayslipWorker(
            calc_path=calc_path,
            master_path=master_path,
            output_dir=output_dir,
            salary_month=salary_month,
            payment_date=payment_date,
            lks_paths=lks_paths,
        )
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.handle_generation_finished)
        self.worker.failed.connect(self.handle_generation_failed)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def handle_generation_finished(self, result) -> None:
        self.last_output_dir = result.output_dir
        self.set_processing_state(False)

        self.append_log(f"Generated Excel payslips: {result.generated_xlsx_count}")
        self.append_log(f"Generated PDF payslips: {result.generated_pdf_count}")
        self.append_log(f"Output folder: {result.output_dir}")
        if result.calculation_workbook_path is not None:
            self.append_log(f"Generated calculation workbook: {result.calculation_workbook_path}")
        if result.claim_summary is not None:
            self.append_log(f"LKS files used: {result.claim_summary.source_files}")
            self.append_log(
                f"CLAIM rows counted: {result.claim_summary.counted_rows} / {result.claim_summary.total_rows}"
            )
            self.append_log(f"CLAIM rows skipped: {result.claim_summary.skipped_rows}")

        if result.warnings:
            self.append_log("")
            self.append_log("Warnings:")
            for warning in result.warnings:
                self.append_log(f"- {warning}")

        if result.pdf_failures:
            self.append_log("")
            self.append_log("PDF export failures:")
            for failure in result.pdf_failures:
                self.append_log(f"- {failure}")

        summary_lines = [
            f"Excel payslips generated: {result.generated_xlsx_count}",
            f"PDF payslips generated: {result.generated_pdf_count}",
            f"Output folder: {result.output_dir}",
        ]
        if result.calculation_workbook_path is not None:
            summary_lines.append(f"Generated calculation workbook: {result.calculation_workbook_path}")
        if result.claim_summary is not None:
            summary_lines.append(f"CLAIM files used: {result.claim_summary.source_files}")
            summary_lines.append(
                f"CLAIM rows counted: {result.claim_summary.counted_rows} / {result.claim_summary.total_rows}"
            )
            summary_lines.append(f"CLAIM rows skipped: {result.claim_summary.skipped_rows}")
        if result.warnings:
            summary_lines.append("")
            summary_lines.append("Warnings:")
            summary_lines.extend(f"- {warning}" for warning in result.warnings)
        if result.pdf_failures:
            summary_lines.append("")
            summary_lines.append("PDF export failures:")
            summary_lines.extend(f"- {failure}" for failure in result.pdf_failures)

        self.summary_text.setPlainText("\n".join(summary_lines))
        QMessageBox.information(
            self,
            "Generation Complete",
            f"Generated {result.generated_xlsx_count} Excel payslips and {result.generated_pdf_count} PDFs.",
        )

    def handle_generation_failed(self, message: str) -> None:
        self.set_processing_state(False)
        self.summary_text.setPlainText("Generation failed.")
        self.append_log(f"Generation failed: {message}")
        QMessageBox.critical(self, "Generation Failed", message)


class PayslipWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"TNB Payslip Generator v{VERSION}")
        self.resize(1160, 760)
        self.setMinimumSize(980, 700)
        self.setCentralWidget(PayslipPanel())


def main() -> int:
    try:
        app = QApplication(sys.argv)
        app.setPalette(apply_app_palette(app.palette()))
        window = PayslipWindow()
        window.showMaximized()
        return app.exec()
    except Exception as exc:  # pragma: no cover - startup diagnostics
        _report_startup_error("Payslip Generator Failed To Start", exc)


if __name__ == "__main__":
    raise SystemExit(main())
