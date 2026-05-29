from __future__ import annotations

import json
import re
import subprocess
import sys
import threading
from datetime import date
from pathlib import Path

from config import DEFAULT_TEMPLATE_PATH
from core.services.payslip_service import (
    DEFAULT_CALC_PATH,
    DEFAULT_LKS_SAMPLE_PATH,
    DEFAULT_MASTER_PATH,
    DEFAULT_OUTPUT_DIR,
    generate_payslips,
)
from main import run_process
from ui_theme import apply_app_palette
from updater import APP_DIR

try:
    from PySide6.QtCore import QObject, QThread, QUrl, Signal, Slot
    from PySide6.QtGui import QAction, QDesktopServices
    from PySide6.QtWebChannel import QWebChannel
    from PySide6.QtWebEngineCore import QWebEngineSettings
    from PySide6.QtWebEngineWidgets import QWebEngineView
    from PySide6.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox
except ImportError as exc:
    raise SystemExit(
        "PySide6 with Qt WebEngine is required to run the modern desktop prototype."
    ) from exc


VERSION_FILE = APP_DIR / "VERSION"
VERSION = VERSION_FILE.read_text(encoding="utf-8").strip() if VERSION_FILE.exists() else "0.0.0"
WEB_UI_DIR = APP_DIR / "web_ui"
INDEX_FILE = WEB_UI_DIR / "index.html"
ANSI_PATTERN = re.compile(r"\x1b\[[0-9;]*m")


class WebLksWorker(QObject):
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


class WebPayslipWorker(QObject):
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


class Bridge(QObject):
    eventEmitted = Signal(str)
    queuedEvent = Signal(str)

    def __init__(self, window: "ModernShellWindow"):
        super().__init__(window)
        self.window = window
        self.lks_thread: QThread | None = None
        self.lks_worker: WebLksWorker | None = None
        self.payslip_thread: QThread | None = None
        self.payslip_worker: WebPayslipWorker | None = None
        self.queuedEvent.connect(self._dispatch_event)

    def _emit(self, payload: dict) -> None:
        self.queuedEvent.emit(json.dumps(payload))

    @Slot(str)
    def _dispatch_event(self, payload_json: str) -> None:
        self.eventEmitted.emit(payload_json)
        try:
            script = (
                "if (window.__tnbDispatchEvent) { "
                f"window.__tnbDispatchEvent({json.dumps(payload_json)});"
                " }"
            )
            self.window.view.page().runJavaScript(script)
        except Exception:
            pass

    @staticmethod
    def _clean_log_message(message: str) -> str:
        return ANSI_PATTERN.sub("", message).rstrip()

    @staticmethod
    def _friendly_lks_error(raw_message: str) -> str:
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

    @Slot(result=str)
    def getInitialState(self) -> str:
        return json.dumps(
            {
                "version": VERSION,
                "supportEmail": "syahmi@nuaim.my",
                "supportPhone": "+60 18 2605 390",
                "defaults": {
                    "lksTemplate": str(Path(DEFAULT_TEMPLATE_PATH).resolve()),
                    "calcPath": str(DEFAULT_CALC_PATH) if DEFAULT_CALC_PATH.exists() else "",
                    "masterPath": str(DEFAULT_MASTER_PATH) if DEFAULT_MASTER_PATH.exists() else "",
                    "outputDir": str(DEFAULT_OUTPUT_DIR),
                    "lksSampleDir": str(DEFAULT_LKS_SAMPLE_PATH.parent),
                },
            }
        )

    @Slot(str, result=str)
    def pickFile(self, kind: str) -> str:
        if kind == "lksInput":
            path, _ = QFileDialog.getOpenFileName(
                self.window,
                "Select Input Excel File",
                str(Path.home()),
                "Excel files (*.xls *.xlsx);;All files (*.*)",
            )
            return path
        if kind == "calc":
            path, _ = QFileDialog.getOpenFileName(
                self.window,
                "Select Calculation Workbook",
                str(DEFAULT_CALC_PATH.parent),
                "Excel files (*.xlsx *.xlsm);;All files (*.*)",
            )
            return path
        if kind == "master":
            path, _ = QFileDialog.getOpenFileName(
                self.window,
                "Select Worker Master Workbook",
                str(DEFAULT_MASTER_PATH.parent),
                "Excel files (*.xlsx *.xlsm);;All files (*.*)",
            )
            return path
        return ""

    @Slot(result=str)
    def pickLksFiles(self) -> str:
        paths, _ = QFileDialog.getOpenFileNames(
            self.window,
            "Select LKS Files",
            str(DEFAULT_LKS_SAMPLE_PATH.parent),
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        return json.dumps(paths)

    @Slot(result=str)
    def pickDirectory(self) -> str:
        folder = QFileDialog.getExistingDirectory(
            self.window,
            "Select Output Folder",
            str(DEFAULT_OUTPUT_DIR),
        )
        return folder

    @Slot(str)
    def checkUpdates(self, _module: str = "") -> None:
        updater_script = APP_DIR / "updater.py"
        try:
            subprocess.Popen([sys.executable, str(updater_script), "--check-only"], cwd=str(APP_DIR))
            self._emit({"type": "toast", "level": "info", "message": "Update check started."})
        except Exception as exc:
            self._emit({"type": "toast", "level": "error", "message": str(exc)})

    @Slot(str)
    def openPath(self, target: str) -> None:
        if not target or not Path(target).exists():
            self._emit({"type": "toast", "level": "error", "message": "The selected file or folder could not be found."})
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(target)):
            self._emit({"type": "toast", "level": "error", "message": f"Could not open: {target}"})

    @Slot(bool)
    def respondAppendConfirmation(self, answer: bool) -> None:
        if self.lks_worker is not None:
            self.lks_worker.set_append_confirmation(answer)

    @Slot(str)
    def startLks(self, payload_json: str) -> None:
        if self.lks_thread is not None:
            self._emit({"type": "toast", "level": "error", "message": "LKS processing is already running."})
            return

        payload = json.loads(payload_json)
        input_path = Path(payload.get("inputPath", "")).expanduser()
        template_path = Path(payload.get("templatePath") or DEFAULT_TEMPLATE_PATH).resolve()

        if not input_path.exists():
            self._emit({"type": "runFailed", "module": "lks", "message": "Select a valid input workbook."})
            return
        if not template_path.exists():
            self._emit({"type": "runFailed", "module": "lks", "message": "The LKS template file could not be found."})
            return

        self.lks_thread = QThread(self.window)
        self.lks_worker = WebLksWorker(input_path.resolve(), template_path)
        self.lks_worker.moveToThread(self.lks_thread)
        self.lks_thread.started.connect(self.lks_worker.run)
        self.lks_worker.log_message.connect(
            lambda message: self._emit(
                {"type": "log", "module": "lks", "message": self._clean_log_message(message)}
            )
        )
        self.lks_worker.status_changed.connect(
            lambda status: self._emit({"type": "status", "module": "lks", "value": status})
        )
        self.lks_worker.append_confirmation_requested.connect(
            lambda existing_count, new_count: self._emit(
                {
                    "type": "appendConfirmation",
                    "module": "lks",
                    "existingCount": existing_count,
                    "newCount": new_count,
                }
            )
        )
        self.lks_worker.finished.connect(self._handle_lks_finished)
        self.lks_worker.failed.connect(self._handle_lks_failed)
        self.lks_worker.finished.connect(self.lks_thread.quit)
        self.lks_worker.failed.connect(self.lks_thread.quit)
        self.lks_thread.finished.connect(self.lks_worker.deleteLater)
        self.lks_thread.finished.connect(self.lks_thread.deleteLater)
        self._emit({"type": "runStarted", "module": "lks"})
        self.lks_thread.start()

    @Slot(object)
    def _handle_lks_finished(self, result: dict) -> None:
        self.lks_worker = None
        self.lks_thread = None
        self._emit(
            {
                "type": "runCompleted",
                "module": "lks",
                "result": {
                    "aborted": result.get("aborted", False),
                    "outputPath": result.get("output_path"),
                    "generatedInputPath": result.get("generated_input_path"),
                    "summary": result.get("summary", {}),
                    "trasByDate": result.get("tras_by_date", {}),
                },
            }
        )

    @Slot(str)
    def _handle_lks_failed(self, message: str) -> None:
        self.lks_worker = None
        self.lks_thread = None
        self._emit(
            {
                "type": "runFailed",
                "module": "lks",
                "message": self._friendly_lks_error(message),
            }
        )

    @Slot(str)
    def startPayslip(self, payload_json: str) -> None:
        if self.payslip_thread is not None:
            self._emit({"type": "toast", "level": "error", "message": "Payslip generation is already running."})
            return

        payload = json.loads(payload_json)
        calc_path = Path(payload.get("calcPath", "")).expanduser()
        master_path = Path(payload.get("masterPath", "")).expanduser()
        output_dir = Path(payload.get("outputDir", "")).expanduser() if payload.get("outputDir") else DEFAULT_OUTPUT_DIR
        salary_month = payload.get("salaryMonth", "").strip()
        payment_date = date.fromisoformat(payload.get("paymentDate"))
        lks_paths = [Path(value).expanduser() for value in payload.get("lksPaths", []) if value]

        if not calc_path.exists():
            self._emit({"type": "runFailed", "module": "payslip", "message": "Select a valid calculation workbook."})
            return
        if not master_path.exists():
            self._emit({"type": "runFailed", "module": "payslip", "message": "Select a valid worker master workbook."})
            return
        if not salary_month:
            self._emit({"type": "runFailed", "module": "payslip", "message": "Enter the salary month label."})
            return
        for path in lks_paths:
            if not path.exists():
                self._emit(
                    {
                        "type": "runFailed",
                        "module": "payslip",
                        "message": f"LKS file not found: {path}",
                    }
                )
                return

        self.payslip_thread = QThread(self.window)
        self.payslip_worker = WebPayslipWorker(
            calc_path=calc_path.resolve(),
            master_path=master_path.resolve(),
            output_dir=output_dir,
            salary_month=salary_month,
            payment_date=payment_date,
            lks_paths=[path.resolve() for path in lks_paths],
        )
        self.payslip_worker.moveToThread(self.payslip_thread)
        self.payslip_thread.started.connect(self.payslip_worker.run)
        self.payslip_worker.log_message.connect(
            lambda message: self._emit(
                {"type": "log", "module": "payslip", "message": self._clean_log_message(message)}
            )
        )
        self.payslip_worker.finished.connect(self._handle_payslip_finished)
        self.payslip_worker.failed.connect(self._handle_payslip_failed)
        self.payslip_worker.finished.connect(self.payslip_thread.quit)
        self.payslip_worker.failed.connect(self.payslip_thread.quit)
        self.payslip_thread.finished.connect(self.payslip_worker.deleteLater)
        self.payslip_thread.finished.connect(self.payslip_thread.deleteLater)
        self._emit({"type": "runStarted", "module": "payslip"})
        self.payslip_thread.start()

    @Slot(object)
    def _handle_payslip_finished(self, result) -> None:
        self.payslip_worker = None
        self.payslip_thread = None
        self._emit(
            {
                "type": "runCompleted",
                "module": "payslip",
                "result": {
                    "outputDir": str(result.output_dir),
                    "generatedXlsxCount": result.generated_xlsx_count,
                    "generatedPdfCount": result.generated_pdf_count,
                    "warnings": list(result.warnings),
                    "pdfFailures": list(result.pdf_failures),
                    "calculationWorkbookPath": (
                        str(result.calculation_workbook_path) if result.calculation_workbook_path else ""
                    ),
                    "claimSummary": (
                        {
                            "sourceFiles": result.claim_summary.source_files,
                            "totalRows": result.claim_summary.total_rows,
                            "countedRows": result.claim_summary.counted_rows,
                            "skippedRows": result.claim_summary.skipped_rows,
                        }
                        if result.claim_summary
                        else None
                    ),
                },
            }
        )

    @Slot(str)
    def _handle_payslip_failed(self, message: str) -> None:
        self.payslip_worker = None
        self.payslip_thread = None
        self._emit({"type": "runFailed", "module": "payslip", "message": message})


class ModernShellWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"TNB Automation Workspace v{VERSION}")
        self.resize(1440, 960)
        self.setMinimumSize(1180, 760)

        self.view = QWebEngineView(self)
        settings = self.view.settings()
        settings.setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(QWebEngineSettings.LocalContentCanAccessRemoteUrls, False)
        settings.setAttribute(QWebEngineSettings.ErrorPageEnabled, True)

        self.bridge = Bridge(self)
        self.channel = QWebChannel(self.view.page())
        self.channel.registerObject("bridge", self.bridge)
        self.view.page().setWebChannel(self.channel)
        self.view.setUrl(QUrl.fromLocalFile(str(INDEX_FILE.resolve())))
        self.setCentralWidget(self.view)

        refresh_action = QAction("Reload UI", self)
        refresh_action.triggered.connect(self.view.reload)
        self.addAction(refresh_action)
        refresh_action.setShortcut("Ctrl+R")


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("TNB Automation Workspace")
    app.setOrganizationName("TNB LKS")
    app.setStyle("Fusion")
    app.setPalette(apply_app_palette(app.palette()))

    if not INDEX_FILE.exists():
        QMessageBox.critical(None, "Missing UI Files", f"Could not find:\n{INDEX_FILE}")
        return 1

    window = ModernShellWindow()
    window.showMaximized()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
