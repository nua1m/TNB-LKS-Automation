from __future__ import annotations

from PySide6.QtGui import QColor, QPalette


def apply_app_palette(palette: QPalette) -> QPalette:
    palette.setColor(QPalette.Window, QColor("#f5f7fa"))
    palette.setColor(QPalette.Base, QColor("#ffffff"))
    palette.setColor(QPalette.Button, QColor("#ffffff"))
    return palette


def panel_stylesheet() -> str:
    return """
    QWidget {
        background: #f5f7fa;
        color: #182230;
        font-family: "Inter", "Segoe UI";
        font-size: 10pt;
    }
    QMainWindow {
        background: #f5f7fa;
    }
    QLabel#pageTitle {
        font-size: 21pt;
        font-weight: 700;
        color: #111827;
    }
    QLabel {
        background: transparent;
    }
    QFrame#fileDropArea {
        background: #ffffff;
        border: 1px dashed #c9d4e2;
        border-radius: 12px;
    }
    QFrame#fileDropArea[dragActive="true"] {
        background: #f3f8ff;
        border: 1px dashed #6f9be7;
    }
    QLabel#dropTitle {
        color: #111827;
        font-size: 12pt;
        font-weight: 700;
        background: transparent;
    }
    QLabel#dropText {
        color: #344054;
        font-size: 10pt;
        font-weight: 600;
        background: transparent;
    }
    QLabel#dropHint {
        color: #667085;
        font-size: 9pt;
        background: transparent;
    }
    QGroupBox {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        font-weight: 600;
        margin-top: 12px;
        padding-top: 16px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 16px;
        padding: 0 4px;
        color: #1f2937;
        background: transparent;
    }
    QPushButton#infoButton {
        min-width: 34px;
        max-width: 34px;
        min-height: 34px;
        max-height: 34px;
        background: #ffffff;
        border: 1px solid #d9e2ec;
        border-radius: 10px;
        color: #475467;
        font-size: 12pt;
        font-weight: 700;
        padding: 0px;
    }
    QPushButton#infoButton:hover {
        background: #f8fafc;
    }
    QLineEdit, QPlainTextEdit, QDateEdit {
        background: #ffffff;
        border: 1px solid #d9e2ec;
        border-radius: 10px;
        padding: 10px 12px;
        selection-background-color: #dbeafe;
        selection-color: #111827;
    }
    QLineEdit[readOnly="true"] {
        background: #ffffff;
        color: #344054;
    }
    QLineEdit:focus, QPlainTextEdit:focus, QDateEdit:focus {
        border: 1px solid #7aa2e3;
    }
    QLineEdit::placeholder {
        color: #98a2b3;
    }
    QPushButton {
        background: #ffffff;
        border: 1px solid #d9e2ec;
        border-radius: 10px;
        padding: 10px 14px;
        min-height: 18px;
        font-weight: 600;
        color: #243041;
    }
    QPushButton:hover {
        background: #f8fafc;
    }
    QPushButton:disabled {
        background: #f8fafc;
        color: #9aa5b1;
        border-color: #e5eaf0;
    }
    QPushButton#primaryButton {
        background: #1f7a4f;
        border: 1px solid #1f7a4f;
        color: #ffffff;
    }
    QPushButton#primaryButton:hover {
        background: #1a6844;
    }
    QPlainTextEdit#summaryPane, QPlainTextEdit#logPane {
        background: #ffffff;
        color: #243041;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
    }
    QPlainTextEdit#logPane {
        color: #344054;
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


def tab_stylesheet() -> str:
    return """
    QTabWidget#workspaceTabs::pane {
        border: 1px solid #e2e8f0;
        background: #ffffff;
        border-radius: 12px;
        top: -1px;
    }
    QTabWidget#workspaceTabs QTabBar::tab {
        background: #edf2f7;
        color: #475467;
        border: 1px solid #d9e2ec;
        border-bottom: none;
        padding: 11px 18px;
        margin-right: 8px;
        min-width: 160px;
        border-top-left-radius: 10px;
        border-top-right-radius: 10px;
        font-weight: 600;
    }
    QTabWidget#workspaceTabs QTabBar::tab:selected {
        background: #ffffff;
        color: #111827;
        border-color: #e2e8f0;
    }
    QTabWidget#workspaceTabs QTabBar::tab:hover:!selected {
        background: #f8fafc;
    }
    QTabWidget#workspaceTabs QTabBar {
        left: 18px;
    }
    """
