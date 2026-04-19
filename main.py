# PreCita - Copyright (C) 2026 eucarigo
# Este programa es software libre: puedes redistribuirlo y/o modificarlo 
# bajo los términos de la Licencia Pública General de GNU publicada por 
# la Free Software Foundation, versión 3.

import sys, json, os, sqlite3, base64, re, mimetypes, shutil, zipfile, hashlib, hmac, inspect, queue, socket, threading, urllib.error, urllib.parse, urllib.request, pytz
from datetime import datetime, timedelta, date
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email import encoders
from pathlib import Path
from uuid import uuid4
from cryptography.fernet import Fernet

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
    QPushButton, QLabel, QTextEdit, QTextBrowser, QTableWidget, 
    QTableWidgetItem, QDialog, QLineEdit, QMessageBox, QSystemTrayIcon,
    QMenu, QFormLayout, QFrame, QSizePolicy, QHeaderView, QAbstractItemView,
    QComboBox, QCheckBox, QGroupBox, QInputDialog, QToolButton, QButtonGroup,
    QGridLayout, QWidgetAction, QFileDialog, QListWidget, QScrollArea, QToolTip,
)
from PyQt6.QtCore import (
    Qt, QTimer, pyqtSignal, QObject, QThread, 
    QRegularExpression, QUrl, QPoint, QRect, QEvent,
)
from PyQt6.QtGui import (
    QIcon, QFont, QColor, QAction, QRegularExpressionValidator,
    QTextBlockFormat, QTextCharFormat, QTextCursor, QTextDocument,
    QTextListFormat, QDesktopServices, QShortcut, QKeySequence,
    QPainter, QPen, QBrush, QImage,
)
from PyQt6.QtNetwork import QLocalServer, QLocalSocket
from PyQt6.QtWebEngineWidgets import QWebEngineView

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

def rpath(rel):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, rel)

# ============================================================================
# CONSTANTES, CONFIGURACIÓN, UTILIDADES
# ============================================================================

PRECITA_MASTER_KEY = None # añade aquí tu clave para descifrar el config.bin

SCOPES_CALENDAR = ['https://www.googleapis.com/auth/calendar.readonly']
SCOPES_GMAIL = ['https://www.googleapis.com/auth/gmail.send']
SCOPES = SCOPES_CALENDAR + SCOPES_GMAIL
OAUTH_LOOPBACK_PORT = 8080
OAUTH_EXTERNAL_TIMEOUT_SECONDS = 15

CLIENT_SECRETS = Path(__file__).parent / "config.bin"

PRECITA_LP = Path.home() / '.precita'
DB_PATH = PRECITA_LP / 'precita.db'
CREDENTIALS_PATH = PRECITA_LP / 'token.json'
ATTACHMENTS_DIR = PRECITA_LP / "template_attachments"
DB_ENCRYPTION_CONFIG_PATH = PRECITA_LP / 'db_encryption_config.json'
MAX_TEMPLATE_PAYLOAD_BYTES = 16 * 1024 * 1024
BLOCKED_ATTACHMENT_EXTENSIONS = {
    ".ade", ".adp", ".apk", ".appx", ".bat", ".cab", ".chm", ".cmd", ".com", ".cpl",
    ".dll", ".dmg", ".exe", ".hta", ".iso", ".jar", ".js", ".jse", ".lnk", ".msc",
    ".msi", ".msp", ".pif", ".ps1", ".scr", ".sys", ".vb", ".vbe", ".vbs",
}
ARCHIVE_EXTENSIONS = {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz"}

PRECITA_LP.mkdir(parents=True, exist_ok=True)
ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)

DB_ENCRYPTION_MAGIC = b"PRECITA_DB_ENC_v1"
DB_ENCRYPTION_SALT_BYTES = 16
DB_ENCRYPTION_NONCE_BYTES = 16
DB_ENCRYPTION_HMAC_BYTES = 32
DB_ENCRYPTION_PBKDF2_ITERATIONS = 200_000

RUNTIME_DB_ENCRYPTION_ENABLED = False
RUNTIME_DB_ENCRYPTION_PASSWORD = None

_PHONE_DIGITS_VALIDATOR = QRegularExpressionValidator(QRegularExpression(r"^\d*$"))
_PLAUSIBLE_EMAIL_RE = re.compile(
    r"^[A-Za-z0-9][A-Za-z0-9._%+\-]{0,63}@"
    r"(?:[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?\.)+"
    r"[A-Za-z]{2,63}$"
)
SPANISH_WEEKDAY_NAMES = (
    "lunes", "martes", "miércoles",
    "jueves", "viernes", "sábado", "domingo",
)
SPANISH_MONTH_NAMES = (
    "enero", "febrero", "marzo", "abril", "mayo", 
    "junio", "julio", "agosto", "septiembre", 
    "octubre", "noviembre", "diciembre",
)
SPANISH_MONTH_ABBR = (
    "ene", "feb", "mar", "abr", "may", "jun", 
    "jul", "ago", "sep", "oct", "nov", "dic",
)

SINGLE_INSTANCE_SERVER_NAME = "precita_single_instance_server"

if sys.platform == "win32":
    import winreg

PRECITA_QSS_LIGHT = """
QMainWindow { background-color: #e8edf3; }
QWidget#centralRoot { background-color: #e8edf3; }

QFrame#appHeader {
    background-color: #ffffff;
    border: none;
    border-bottom: 1px solid #cfd8e3;
}
QLabel#brandTitle {
    color: #0c1929;
    font-size: 20px;
    font-weight: 700;
    letter-spacing: -0.3px;
}
QLabel#googleStatusDot {
    min-width: 12px;
    max-width: 12px;
    min-height: 12px;
    max-height: 12px;
    border-radius: 6px;
    border: 1px solid #8aa4bf;
}
QLabel#googleStatusDot[syncState="synced"] {
    background-color: #31c451;
    border-color: #24953d;
}
QLabel#googleStatusDot[syncState="unsynced"] {
    background-color: #e05252;
    border-color: #ab2f2f;
}
QLabel#brandSubtitle {
    color: #5a6b80;
    font-size: 12px;
}
QLabel#sectionTitle {
    color: #2c3e50;
    font-size: 12px;
    font-weight: 600;
}

QFrame#panelCard {
    background-color: #ffffff;
    border: 1px solid #cfd8e3;
    border-radius: 10px;
}

QPushButton {
    background-color: #eef2f7;
    color: #1c2836;
    border: 1px solid #b8c5d4;
    border-radius: 8px;
    padding: 8px 14px;
    font-size: 12px;
    font-weight: 500;
    min-height: 18px;
}
QPushButton:hover {
    background-color: #e2e9f2;
    border-color: #9aacbf;
}
QPushButton:pressed { background-color: #d5dee9; }

QPushButton#primaryButton {
    background-color: #1565a0;
    color: #ffffff;
    border: 1px solid #0d5285;
    font-weight: 600;
}
QPushButton#primaryButton:hover {
    background-color: #125a8f;
    border-color: #0a4670;
}
QPushButton#primaryButton:pressed { background-color: #0f4d7a; }
QPushButton#formatToggleButton:checked {
    background-color: #1565a0;
    color: #ffffff;
    border: 1px solid #0d5285;
}
QPushButton#formatToggleButton:checked:hover {
    background-color: #125a8f;
    border-color: #0a4670;
}

QFrame#calendarToolbar {
    background-color: #ffffff;
    border: 1px solid #d2dbe7;
    border-radius: 10px;
}
QPushButton#calendarNavButton {
    background-color: #ffffff;
    color: #425466;
    border: 1px solid #d2dbe7;
    border-radius: 18px;
    min-width: 34px;
    min-height: 28px;
    padding: 4px 8px;
    font-size: 13px;
    font-weight: 600;
}
QPushButton#calendarNavButton:hover { background-color: #f6f9fe; }
QPushButton#calendarTodayButton {
    background-color: #ffffff;
    color: #1f2d3d;
    border: 1px solid #d2dbe7;
    border-radius: 18px;
    min-height: 28px;
    padding: 4px 14px;
    font-size: 12px;
    font-weight: 600;
}
QPushButton#calendarTodayButton:hover { background-color: #f6f9fe; }
QLabel#calendarPeriodTitle {
    color: #1f2d3d;
    font-size: 17px;
    font-weight: 500;
}
QPushButton#calendarViewButton {
    background-color: #ffffff;
    color: #2f3e4f;
    border: 1px solid #d2dbe7;
    border-radius: 15px;
    min-width: 78px;
    min-height: 30px;
    padding: 5px 14px;
    font-size: 12px;
    font-weight: 500;
}
QPushButton#calendarViewButton:hover { background-color: #f6f9fe; }
QPushButton#calendarViewButton:checked {
    background-color: #d3e3fd;
    border: 1px solid #a8c7fa;
    color: #0842a0;
    font-weight: 600;
}

QToolButton#headerMenuButton {
    background-color: #eef2f7;
    color: #1c2836;
    border: 1px solid #b8c5d4;
    border-radius: 8px;
    padding: 6px 10px;
    font-size: 16px;
    font-weight: 700;
    min-width: 32px;
    min-height: 28px;
}
QToolButton#headerMenuButton:hover {
    background-color: #e2e9f2;
    border-color: #9aacbf;
}
QToolButton#headerMenuButton:pressed {
    background-color: #d5dee9;
}
QToolButton#headerMenuButton::menu-indicator {
    image: none;
    width: 0px;
}

QTableWidget {
    background-color: #ffffff;
    alternate-background-color: #f6f8fb;
    gridline-color: #e3e9f0;
    border: none;
    border-radius: 8px;
    font-size: 12px;
    color: #1c2836;
    selection-background-color: #c5ddf0;
    selection-color: #0c1929;
}
QTableWidget::item { padding: 6px 10px; border: none; color: #1c2836; }
QTableWidget::item:selected { background-color: #c5ddf0; }

QHeaderView::section {
    background-color: #eef2f7;
    color: #4a5f73;
    padding: 9px 11px;
    border: none;
    border-bottom: 1px solid #cfd8e3;
    border-right: 1px solid #e3e9f0;
    font-weight: 600;
    font-size: 11px;
}
QHeaderView::section:last { border-right: none; }

QTextEdit {
    background-color: #ffffff;
    border: 1px solid #cfd8e3;
    border-radius: 8px;
    padding: 8px 10px;
    font-size: 12px;
    color: #1c2836;
}
QTextEdit#activityLog {
    background-color: #f4f7fa;
    color: #4a5f73;
    font-family: "Cascadia Mono", "Consolas", "Lucida Console", monospace;
    font-size: 11px;
}

QLineEdit {
    background-color: #ffffff;
    border: 1px solid #cfd8e3;
    border-radius: 8px;
    padding: 7px 11px;
    font-size: 12px;
    color: #1c2836;
}
QLineEdit:focus { border: 1px solid #1565a0; }

QAbstractItemView QLineEdit {
    padding: 3px 8px;
    min-height: 24px;
    border-radius: 6px;
}

QDialog { background-color: #e8edf3; }
QScrollArea#settingsScrollArea,
QWidget#settingsContent {
    background-color: #e8edf3;
}
QDialog QLabel { color: #2c3e50; }
QDialog QLabel#dialogHint {
    color: #5a6b80;
    font-size: 11px;
    padding: 8px 0;
}

QScrollBar:vertical {
    background: #f0f3f7;
    width: 11px;
    border-radius: 5px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #b9c5d3;
    min-height: 24px;
    border-radius: 5px;
    margin: 2px;
}
QScrollBar::handle:vertical:hover { background: #9aaab8; }
QScrollBar:horizontal {
    background: #f0f3f7;
    height: 11px;
    border-radius: 5px;
}
QScrollBar::handle:horizontal {
    background: #b9c5d3;
    border-radius: 5px;
    margin: 2px;
}

QMenu {
    background-color: #ffffff;
    border: 1px solid #cfd8e3;
    border-radius: 8px;
    padding: 4px;
}
QMenu::item {
    padding: 8px 24px 8px 12px;
    border-radius: 6px;
    color: #1c2836;
}
QWidget#syncMenuActionWidget {
    border-radius: 6px;
}
QWidget#syncMenuActionWidget:hover {
    background-color: #d9ebf7;
}
QWidget#launchPendingMenuActionWidget {
    border-radius: 6px;
}
QWidget#launchPendingMenuActionWidget:hover {
    background-color: #d9ebf7;
}
QPushButton#launchPendingMenuActionButton {
    background: transparent;
    border: none;
    color: #1c2836;
    text-align: left;
    padding: 0px;
    font-size: 13px;
}
QPushButton#launchPendingMenuActionButton:hover {
    color: #1c2836;
}
QPushButton#launchPendingMenuActionButton:disabled {
    color: #9aa7b5;
}
QPushButton#syncMenuActionButton {
    background: transparent;
    border: none;
    color: #1c2836;
    text-align: left;
    padding: 0px;
    font-size: 13px;
}
QPushButton#syncMenuActionButton:hover {
    color: #1c2836;
}
QWidget#logoutMenuActionWidget {
    border-radius: 6px;
}
QWidget#logoutMenuActionWidget:hover {
    background-color: #fce8e6;
}
QPushButton#logoutMenuActionButton {
    background: transparent;
    border: none;
    color: #1c2836;
    text-align: left;
    padding: 0px;
    font-size: 13px;
}
QPushButton#logoutMenuActionButton:hover {
}
QPushButton#logoutMenuActionButton:disabled {
    color: #9aa7b5;
}
QPushButton#inlineImageButton[softDisabled="true"] {
    background-color: #eef2f6;
    color: #9aa7b5;
    border-color: #d5dde6;
}
QPushButton#inlineImageButton[softDisabled="true"]:hover {
    background-color: #eef2f6;
    color: #9aa7b5;
    border-color: #d5dde6;
}
QWidget#logoutMenuActionWidget:disabled {
    background-color: #eef2f6;
}
QWidget#logoutMenuActionWidget:disabled:hover {
    background-color: #eef2f6;
}
QWidget#launchPendingMenuActionWidget:disabled {
    background-color: #eef2f6;
}
QWidget#launchPendingMenuActionWidget:disabled:hover {
    background-color: #eef2f6;
}
QMenu::item:selected {
    background-color: #d9ebf7;
    color: #0d5285;
}
QMenu::item#precitaMenuDelete:selected,
QMenu::item[precitaDestructive="true"]:selected {
    background-color: #fce8e6;
    color: #9f1818;
    font-weight: 600;
}
QMenu::separator {
    height: 1px;
    margin: 6px 8px;
    background-color: #d7e0ea;
}

QMessageBox { background-color: #ffffff; }
QMessageBox QLabel { color: #2c3e50; font-size: 12px; }
QMessageBox QPushButton { min-width: 76px; }

QComboBox, QCheckBox, QGroupBox { color: #2c3e50; font-size: 12px; }
QComboBox {
    combobox-popup: 0;
    background-color: #ffffff;
    border: 1px solid #cfd8e3;
    border-radius: 8px;
    padding: 6px 10px;
    min-height: 18px;
    color: #1c2836;
}
QComboBox::drop-down { border: none; width: 22px; }
QComboBox QAbstractItemView {
    background-color: #ffffff;
    color: #1c2836;
    selection-background-color: #d9ebf7;
    selection-color: #0d5285;
    outline: 0;
    border: 1px solid #cfd8e3;
    border-radius: 6px;
    padding: 4px;
}
QGroupBox {
    font-weight: 600;
    border: 1px solid #cfd8e3;
    border-radius: 10px;
    margin-top: 10px;
    padding: 14px 12px 12px 12px;
    background-color: #ffffff;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
}
"""

PRECITA_QSS_DARK = """
QMainWindow { background-color: #1a1f26; }
QWidget#centralRoot { background-color: #1a1f26; }

QFrame#appHeader {
    background-color: #252b35;
    border: none;
    border-bottom: 1px solid #3d4754;
}
QLabel#brandTitle {
    color: #e8edf3;
    font-size: 20px;
    font-weight: 700;
    letter-spacing: -0.3px;
}
QLabel#googleStatusDot {
    min-width: 12px;
    max-width: 12px;
    min-height: 12px;
    max-height: 12px;
    border-radius: 6px;
    border: 1px solid #6c8098;
}
QLabel#googleStatusDot[syncState="synced"] {
    background-color: #39d65b;
    border-color: #2aa647;
}
QLabel#googleStatusDot[syncState="unsynced"] {
    background-color: #ff6262;
    border-color: #c44a4a;
}
QLabel#brandSubtitle {
    color: #9aacbf;
    font-size: 12px;
}
QLabel#sectionTitle {
    color: #c5d3e0;
    font-size: 12px;
    font-weight: 600;
}

QFrame#panelCard {
    background-color: #252b35;
    border: 1px solid #3d4754;
    border-radius: 10px;
}

QPushButton {
    background-color: #2f3845;
    color: #e8edf3;
    border: 1px solid #4a5a6e;
    border-radius: 8px;
    padding: 8px 14px;
    font-size: 12px;
    font-weight: 500;
    min-height: 18px;
}
QPushButton:hover {
    background-color: #3a4555;
    border-color: #5c6f86;
}
QPushButton:pressed { background-color: #252e3a; }

QPushButton#primaryButton {
    background-color: #2d7ab8;
    color: #ffffff;
    border: 1px solid #1f5f94;
    font-weight: 600;
}
QPushButton#primaryButton:hover {
    background-color: #2680c4;
    border-color: #1a5080;
}
QPushButton#primaryButton:pressed { background-color: #1f6499; }
QPushButton#formatToggleButton:checked {
    background-color: #2d7ab8;
    color: #ffffff;
    border: 1px solid #1f5f94;
}
QPushButton#formatToggleButton:checked:hover {
    background-color: #2680c4;
    border-color: #1a5080;
}

QFrame#calendarToolbar {
    background-color: #252b35;
    border: 1px solid #3d4754;
    border-radius: 10px;
}
QPushButton#calendarNavButton {
    background-color: #252b35;
    color: #d7e4f2;
    border: 1px solid #4a5a6e;
    border-radius: 18px;
    min-width: 34px;
    min-height: 28px;
    padding: 4px 8px;
    font-size: 13px;
    font-weight: 600;
}
QPushButton#calendarNavButton:hover { background-color: #2f3845; }
QPushButton#calendarTodayButton {
    background-color: #252b35;
    color: #e8edf3;
    border: 1px solid #4a5a6e;
    border-radius: 18px;
    min-height: 28px;
    padding: 4px 14px;
    font-size: 12px;
    font-weight: 600;
}
QPushButton#calendarTodayButton:hover { background-color: #2f3845; }
QLabel#calendarPeriodTitle {
    color: #e8edf3;
    font-size: 17px;
    font-weight: 500;
}
QPushButton#calendarViewButton {
    background-color: #252b35;
    color: #d7e4f2;
    border: 1px solid #4a5a6e;
    border-radius: 15px;
    min-width: 78px;
    min-height: 30px;
    padding: 5px 14px;
    font-size: 12px;
    font-weight: 500;
}
QPushButton#calendarViewButton:hover { background-color: #2f3845; }
QPushButton#calendarViewButton:checked {
    background-color: #315a8a;
    border: 1px solid #4a7fb8;
    color: #ffffff;
    font-weight: 600;
}

QToolButton#headerMenuButton {
    background-color: #1c2836;
    color: #ffffff;
    border: 1px solid #0c1929;
    border-radius: 8px;
    padding: 6px 10px;
    font-size: 16px;
    font-weight: 700;
    min-width: 32px;
    min-height: 28px;
}
QToolButton#headerMenuButton:hover {
    background-color: #27384b;
    border-color: #162333;
}
QToolButton#headerMenuButton:pressed {
    background-color: #13202e;
}
QToolButton#headerMenuButton::menu-indicator {
    image: none;
    width: 0px;
}

QTableWidget {
    background-color: #1e242d;
    alternate-background-color: #252b35;
    gridline-color: #3d4754;
    border: none;
    border-radius: 8px;
    font-size: 12px;
    color: #e8edf3;
    selection-background-color: #2d5a87;
    selection-color: #ffffff;
}
QTableWidget::item { padding: 6px 10px; border: none; }
QTableWidget::item:selected { background-color: #2d5a87; }

QHeaderView::section {
    background-color: #2f3845;
    color: #b8c9dc;
    padding: 9px 11px;
    border: none;
    border-bottom: 1px solid #3d4754;
    border-right: 1px solid #3d4754;
    font-weight: 600;
    font-size: 11px;
}
QHeaderView::section:last { border-right: none; }

QTextEdit {
    background-color: #1e242d;
    border: 1px solid #3d4754;
    border-radius: 8px;
    padding: 8px 10px;
    font-size: 12px;
    color: #e8edf3;
}
QTextEdit#activityLog {
    background-color: #1a1f26;
    color: #9aacbf;
    font-family: "Cascadia Mono", "Consolas", "Lucida Console", monospace;
    font-size: 11px;
}

QLineEdit {
    background-color: #1e242d;
    border: 1px solid #3d4754;
    border-radius: 8px;
    padding: 7px 11px;
    font-size: 12px;
    color: #e8edf3;
}
QLineEdit:focus { border: 1px solid #2d7ab8; }

QAbstractItemView QLineEdit {
    padding: 3px 8px;
    min-height: 24px;
    border-radius: 6px;
}

QDialog { background-color: #1a1f26; }
QScrollArea#settingsScrollArea,
QWidget#settingsContent {
    background-color: #1a1f26;
}
QDialog QLabel { color: #c5d3e0; }
QDialog QLabel#dialogHint {
    color: #9aacbf;
    font-size: 11px;
    padding: 8px 0;
}

QScrollBar:vertical {
    background: #252b35;
    width: 11px;
    border-radius: 5px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #4a5a6e;
    min-height: 24px;
    border-radius: 5px;
    margin: 2px;
}
QScrollBar::handle:vertical:hover { background: #5c6f86; }
QScrollBar:horizontal {
    background: #252b35;
    height: 11px;
    border-radius: 5px;
}
QScrollBar::handle:horizontal {
    background: #4a5a6e;
    border-radius: 5px;
    margin: 2px;
}

QMenu {
    background-color: #252b35;
    border: 1px solid #3d4754;
    border-radius: 8px;
    padding: 4px;
}
QMenu::item {
    padding: 8px 24px 8px 12px;
    border-radius: 6px;
    color: #e8edf3;
}
QWidget#syncMenuActionWidget {
    border-radius: 6px;
}
QWidget#syncMenuActionWidget:hover {
    background-color: #2d5a87;
}
QWidget#launchPendingMenuActionWidget {
    border-radius: 6px;
}
QWidget#launchPendingMenuActionWidget:hover {
    background-color: #2d5a87;
}
QPushButton#launchPendingMenuActionButton {
    background: transparent;
    border: none;
    color: #e8edf3;
    text-align: left;
    padding: 0px;
    font-size: 13px;
}
QPushButton#launchPendingMenuActionButton:hover {
    color: #e8edf3;
}
QPushButton#launchPendingMenuActionButton:disabled {
    color: #6f7e90;
}
QPushButton#syncMenuActionButton {
    background: transparent;
    border: none;
    color: #e8edf3;
    text-align: left;
    padding: 0px;
    font-size: 13px;
}
QPushButton#syncMenuActionButton:hover {
    color: #e8edf3;
}
QWidget#logoutMenuActionWidget {
    border-radius: 6px;
}
QWidget#logoutMenuActionWidget:hover {
    background-color: #4a2826;
}
QPushButton#logoutMenuActionButton {
    background: transparent;
    border: none;
    color: #e8edf3;
    text-align: left;
    padding: 0px;
    font-size: 13px;
}
QPushButton#logoutMenuActionButton:hover {
}
QPushButton#logoutMenuActionButton:disabled {
    color: #6f7e90;
}
QPushButton#inlineImageButton[softDisabled="true"] {
    background-color: #2b313b;
    color: #6f7e90;
    border-color: #3d4754;
}
QPushButton#inlineImageButton[softDisabled="true"]:hover {
    background-color: #2b313b;
    color: #6f7e90;
    border-color: #3d4754;
}
QWidget#logoutMenuActionWidget:disabled {
    background-color: #2b313b;
}
QWidget#logoutMenuActionWidget:disabled:hover {
    background-color: #2b313b;
}
QWidget#launchPendingMenuActionWidget:disabled {
    background-color: #2b313b;
}
QWidget#launchPendingMenuActionWidget:disabled:hover {
    background-color: #2b313b;
}
QMenu::item:selected {
    background-color: #2d5a87;
    color: #ffffff;
}
QMenu::item#precitaMenuDelete:selected,
QMenu::item[precitaDestructive="true"]:selected {
    background-color: #4a2826;
    color: #ffb3ae;
    font-weight: 600;
}
QMenu::separator {
    height: 1px;
    margin: 6px 8px;
    background-color: #3d4754;
}

QMessageBox { background-color: #252b35; }
QMessageBox QLabel { color: #c5d3e0; font-size: 12px; }
QMessageBox QPushButton { min-width: 76px; }

QComboBox, QCheckBox, QGroupBox { color: #c5d3e0; font-size: 12px; }
QComboBox {
    combobox-popup: 0;
    background-color: #1e242d;
    border: 1px solid #3d4754;
    border-radius: 8px;
    padding: 6px 10px;
    min-height: 18px;
    color: #e8edf3;
}
QComboBox::drop-down { border: none; width: 22px; }
QComboBox QAbstractItemView {
    background-color: #1e242d;
    color: #e8edf3;
    selection-background-color: #2d5a87;
    selection-color: #ffffff;
    outline: 0;
    border: 1px solid #3d4754;
    border-radius: 6px;
    padding: 4px;
}
QGroupBox {
    font-weight: 600;
    border: 1px solid #3d4754;
    border-radius: 10px;
    margin-top: 10px;
    padding: 14px 12px 12px 12px;
    background-color: #252b35;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
}
"""


def get_version():
    try:
        with open(Path(__file__).parent / 'VERSION', "r") as f:
            return f.read().strip()
    except:
        return "0.36.8 (alpha)"

# ============================================================================
# INICIALIZACIÓN DE BASE DE DATOS
# ============================================================================

def contact_full_name(first_name, last_name):
    """Nombre para mostrar y correos: nombre + apellido(s), sin espacios sobrantes."""
    parts = []
    if first_name and str(first_name).strip():
        parts.append(str(first_name).strip())
    if last_name and str(last_name).strip():
        parts.append(str(last_name).strip())
    return " ".join(parts)


def is_plausible_email(email):
    """Valida formato de correo plausible (no garantiza existencia real)."""
    if not email:
        return False
    cleaned = email.strip()
    if " " in cleaned or ".." in cleaned:
        return False
    return bool(_PLAUSIBLE_EMAIL_RE.fullmatch(cleaned))


def is_known_mail(email):
    # Lista de dominios de confianza
    DOMINIOS_TRUSTED = {
        "gmail.com", "outlook.com", "hotmail.com", "live.com",
        "icloud.com", "yahoo.com", "proton.me", "protonmail.com",
        "zoho.com", "gmx.com", "tuta.io"
    }

    try:
        dominio = email.split('@')[-1].lower()
        return dominio in DOMINIOS_TRUSTED
    except Exception:
        return False


def _migrate_contacts_name_to_first_last(conn):
    """Pasa la columna única `name` a `first_name` + `last_name` (BD antigua)."""
    cursor = conn.cursor()
    cursor.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='contacts'"
    )
    if not cursor.fetchone():
        return
    cursor.execute("PRAGMA table_info(contacts)")
    columns = {row[1] for row in cursor.fetchall()}
    if "first_name" in columns:
        return
    if "name" not in columns:
        return
    cursor.execute("PRAGMA foreign_keys=OFF")
    cursor.execute(
        """
        CREATE TABLE contacts_migrated (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL DEFAULT '',
            email TEXT NOT NULL UNIQUE,
            phone TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cursor.execute(
        """
        INSERT INTO contacts_migrated (id, first_name, last_name, email, phone, created_at)
        SELECT id, name, '', email, phone, created_at FROM contacts
        """
    )
    cursor.execute("DROP TABLE contacts")
    cursor.execute("ALTER TABLE contacts_migrated RENAME TO contacts")
    cursor.execute("PRAGMA foreign_keys=ON")


def _migrate_appointments_add_source_calendar_id(conn):
    """Agrega source_calendar_id para evitar conflictos entre calendarios."""
    cursor = conn.cursor()
    cursor.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='appointments'"
    )
    if not cursor.fetchone():
        return
    cursor.execute("PRAGMA table_info(appointments)")
    columns = {row[1] for row in cursor.fetchall()}
    if "source_calendar_id" not in columns:
        cursor.execute(
            "ALTER TABLE appointments ADD COLUMN source_calendar_id TEXT DEFAULT 'primary'"
        )
    cursor.execute(
        "UPDATE appointments SET source_calendar_id = 'primary' "
        "WHERE source_calendar_id IS NULL OR TRIM(source_calendar_id) = ''"
    )
    cursor.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='appointments'"
    )
    row = cursor.fetchone()
    ddl = (row[0] or "").upper() if row and row[0] else ""
    needs_unique_migration = "CALENDAR_EVENT_ID TEXT UNIQUE" in ddl
    if needs_unique_migration:
        cursor.execute("PRAGMA foreign_keys=OFF")
        cursor.execute(
            """
            CREATE TABLE appointments_migrated (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                calendar_event_id TEXT,
                source_calendar_id TEXT DEFAULT 'primary',
                contact_id INTEGER,
                event_title TEXT,
                event_description TEXT,
                event_date TIMESTAMP,
                reminder_sent INTEGER DEFAULT 0,
                email_sent_at TIMESTAMP,
                FOREIGN KEY (contact_id) REFERENCES contacts(id)
            )
            """
        )
        cursor.execute(
            """
            INSERT INTO appointments_migrated (
                id,
                calendar_event_id,
                source_calendar_id,
                contact_id,
                event_title,
                event_description,
                event_date,
                reminder_sent,
                email_sent_at
            )
            SELECT
                id,
                calendar_event_id,
                source_calendar_id,
                contact_id,
                event_title,
                event_description,
                event_date,
                reminder_sent,
                email_sent_at
            FROM appointments
            """
        )
        cursor.execute("DROP TABLE appointments")
        cursor.execute("ALTER TABLE appointments_migrated RENAME TO appointments")
        cursor.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_appointments_calendar_event_per_calendar "
            "ON appointments(calendar_event_id, source_calendar_id)"
        )
        cursor.execute("PRAGMA foreign_keys=ON")
    else:
        cursor.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_appointments_calendar_event_per_calendar "
            "ON appointments(calendar_event_id, source_calendar_id)"
        )


def init_database():
    """Inicializa la base de datos SQLite con las tablas necesarias."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Tabla de contactos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL DEFAULT '',
            email TEXT NOT NULL UNIQUE,
            phone TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    _migrate_contacts_name_to_first_last(conn)
    
    # Tabla de citas
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS appointments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            calendar_event_id TEXT,
            source_calendar_id TEXT DEFAULT 'primary',
            contact_id INTEGER,
            event_title TEXT,
            event_description TEXT,
            event_date TIMESTAMP,
            reminder_sent INTEGER DEFAULT 0,
            email_sent_at TIMESTAMP,
            FOREIGN KEY (contact_id) REFERENCES contacts(id)
        )
    ''')
    _migrate_appointments_add_source_calendar_id(conn)
    
    # Tabla de configuración
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')
    
    # Insertar configuración por defecto si no existe
    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "email_template"')
    if cursor.fetchone()[0] == 0:
        default_template = (
            "Hola {nombre_citado} {apellidos_citado},\n\n"
            "Este es un recordatorio de su cita para el {dia_semana} {fecha_cita} a las {hora_cita}.\n\n"
            "Datos de la cita:\n"
            "- Nombre: {nombre_citado} {apellidos_citado}\n"
            "- Correo: {correo_citado}\n"
            "- Telefono: {tlf_citado}\n"
            "- Fecha: {fecha_cita}\n"
            "- Hora: {hora_cita}\n\n"
            "Si no puede asistir, por favor contacte con nosotros con anticipación.\n\n"
            "Saludos cordiales."
        )
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('email_template', default_template)
        )

    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "email_subject_template"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('email_subject_template', 'Recordatorio de cita - {nombre_citado} {apellidos_citado}')
        )

    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "email_template_format"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('email_template_format', 'plain')
        )

    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "email_template_attachments"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('email_template_attachments', '[]')
        )
    
    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "gmail_enabled"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('gmail_enabled', '0')
        )

    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "theme"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('theme', 'dark')
        )
    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "display_scale_percent"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('display_scale_percent', str(DEFAULT_DISPLAY_SCALE_PERCENT))
        )

    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "reminder_interval_sec"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('reminder_interval_sec', '300')
        )

    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "calendar_sync_days"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('calendar_sync_days', '15')
        )
    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "google_calendar_id"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('google_calendar_id', 'primary')
        )
    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "google_calendar_name"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('google_calendar_name', 'Calendario principal')
        )

    cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "windows_notifications_enabled"')
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?)',
            ('windows_notifications_enabled', DEFAULT_WINDOWS_NOTIFICATIONS)
        )

    if sys.platform == "win32":
        cursor.execute('SELECT COUNT(*) FROM settings WHERE key = "windows_startup"')
        if cursor.fetchone()[0] == 0:
            cursor.execute(
                'INSERT INTO settings (key, value) VALUES (?, ?)',
                ('windows_startup', DEFAULT_WINDOWS_STARTUP)
            )
    
    conn.commit()
    conn.close()


def get_setting(key, default=None):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT value FROM settings WHERE key = ?', (key,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else default


def get_calendar_sync_days():
    raw_days = get_setting(
        "calendar_sync_days",
        str(DEFAULT_CALENDAR_SYNC_DAYS),
    )
    try:
        days = int(raw_days)
    except (TypeError, ValueError):
        return DEFAULT_CALENDAR_SYNC_DAYS
    if MIN_CALENDAR_SYNC_DAYS <= days <= MAX_CALENDAR_SYNC_DAYS:
        return days
    return DEFAULT_CALENDAR_SYNC_DAYS


def get_selected_google_calendar():
    """Retorna el calendario de Google elegido por el usuario."""
    calendar_id = (
        get_setting("google_calendar_id", "primary") or "primary"
    ).strip() or "primary"
    calendar_name = (
        get_setting("google_calendar_name", "Calendario principal")
        or "Calendario principal"
    ).strip() or "Calendario principal"
    return calendar_id, calendar_name


def list_google_calendars(service):
    """Lista calendarios visibles de la cuenta autenticada."""
    calendars = []
    page_token = None
    while True:
        response = service.calendarList().list(pageToken=page_token).execute()
        for item in response.get("items", []):
            cal_id = (item.get("id") or "").strip()
            if not cal_id:
                continue
            cal_name = (item.get("summary") or cal_id).strip()
            is_primary = bool(item.get("primary"))
            calendars.append(
                {
                    "id": cal_id,
                    "name": cal_name,
                    "is_primary": is_primary,
                }
            )
        page_token = response.get("nextPageToken")
        if not page_token:
            break
    calendars.sort(key=lambda cal: (not cal["is_primary"], cal["name"].lower()))
    return calendars


def prompt_google_calendar_selection(parent, service):
    """Muestra selector de calendario y guarda el elegido."""
    calendars = list_google_calendars(service)
    if not calendars:
        raise RuntimeError(
            "No se encontraron calendarios visibles en esta cuenta de Google."
        )

    options = []
    option_to_calendar = {}
    for cal in calendars:
        suffix = " (principal)" if cal["is_primary"] else ""
        label = f"{cal['name']}{suffix} — {cal['id']}"
        options.append(label)
        option_to_calendar[label] = cal

    current_id, _ = get_selected_google_calendar()
    default_index = 0
    for idx, cal in enumerate(calendars):
        if cal["id"] == current_id:
            default_index = idx
            break

    selected_label, ok = QInputDialog.getItem(
        parent,
        "Seleccionar calendario",
        "Elija el calendario que PreCita puede sincronizar:",
        options,
        default_index,
        False,
    )
    if not ok:
        return None
    selected = option_to_calendar[selected_label]
    set_setting("google_calendar_id", selected["id"])
    set_setting("google_calendar_name", selected["name"])
    return selected


def set_setting(key, value):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'INSERT INTO settings (key, value) VALUES (?, ?) '
        'ON CONFLICT(key) DO UPDATE SET value = excluded.value',
        (key, value)
    )
    conn.commit()
    conn.close()


def _load_db_encryption_config():
    default_cfg = {"enabled": False}
    if not DB_ENCRYPTION_CONFIG_PATH.exists():
        return default_cfg
    try:
        payload = json.loads(DB_ENCRYPTION_CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default_cfg
    return {"enabled": bool(payload.get("enabled", False))}


def _save_db_encryption_config(enabled):
    payload = {"enabled": bool(enabled)}
    DB_ENCRYPTION_CONFIG_PATH.write_text(
        json.dumps(payload, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _db_encryption_enabled_in_config():
    return _load_db_encryption_config().get("enabled", False)


def _is_db_file_encrypted(path: Path):
    if not path.exists():
        return False
    try:
        with path.open("rb") as f:
            return f.read(len(DB_ENCRYPTION_MAGIC)) == DB_ENCRYPTION_MAGIC
    except OSError:
        return False


def _derive_db_key(password, salt, iterations=DB_ENCRYPTION_PBKDF2_ITERATIONS):
    if not isinstance(password, str):
        raise ValueError("La clave debe ser una cadena.")
    normalized = password.strip()
    if not normalized:
        raise ValueError("La clave no puede estar vacia.")
    return hashlib.pbkdf2_hmac(
        "sha256",
        normalized.encode("utf-8"),
        salt,
        int(iterations),
        dklen=32,
    )


def _xor_stream_cipher(raw_bytes, key, nonce):
    output = bytearray(len(raw_bytes))
    offset = 0
    counter = 0
    while offset < len(raw_bytes):
        stream_block = hashlib.blake2b(
            key + nonce + counter.to_bytes(8, "big"),
            digest_size=32,
        ).digest()
        chunk = raw_bytes[offset:offset + 32]
        for idx, b in enumerate(chunk):
            output[offset + idx] = b ^ stream_block[idx]
        offset += len(chunk)
        counter += 1
    return bytes(output)


def _encrypt_db_payload(plain_bytes, password):
    salt = os.urandom(DB_ENCRYPTION_SALT_BYTES)
    nonce = os.urandom(DB_ENCRYPTION_NONCE_BYTES)
    key = _derive_db_key(password, salt)
    cipher_bytes = _xor_stream_cipher(plain_bytes, key, nonce)
    hmac_value = hashlib.sha256(key + nonce + cipher_bytes).digest()
    return (
        DB_ENCRYPTION_MAGIC
        + salt
        + nonce
        + int(DB_ENCRYPTION_PBKDF2_ITERATIONS).to_bytes(4, "big")
        + hmac_value
        + cipher_bytes
    )


def _decrypt_db_payload(encrypted_bytes, password):
    magic_len = len(DB_ENCRYPTION_MAGIC)
    fixed_header_len = (
        magic_len
        + DB_ENCRYPTION_SALT_BYTES
        + DB_ENCRYPTION_NONCE_BYTES
        + 4
        + DB_ENCRYPTION_HMAC_BYTES
    )
    if len(encrypted_bytes) < fixed_header_len:
        raise ValueError("El fichero cifrado no tiene un formato valido.")
    if encrypted_bytes[:magic_len] != DB_ENCRYPTION_MAGIC:
        raise ValueError("El fichero no corresponde al formato cifrado de PreCita.")

    offset = magic_len
    salt = encrypted_bytes[offset:offset + DB_ENCRYPTION_SALT_BYTES]
    offset += DB_ENCRYPTION_SALT_BYTES
    nonce = encrypted_bytes[offset:offset + DB_ENCRYPTION_NONCE_BYTES]
    offset += DB_ENCRYPTION_NONCE_BYTES
    iterations = int.from_bytes(encrypted_bytes[offset:offset + 4], "big")
    offset += 4
    received_hmac = encrypted_bytes[offset:offset + DB_ENCRYPTION_HMAC_BYTES]
    offset += DB_ENCRYPTION_HMAC_BYTES
    cipher_bytes = encrypted_bytes[offset:]

    key = _derive_db_key(password, salt, max(1, iterations))
    expected_hmac = hashlib.sha256(key + nonce + cipher_bytes).digest()
    if not hmac.compare_digest(received_hmac, expected_hmac):
        raise ValueError("Clave incorrecta o fichero cifrado manipulado.")

    plain_bytes = _xor_stream_cipher(cipher_bytes, key, nonce)
    if not plain_bytes.startswith(b"SQLite format 3\x00"):
        raise ValueError("No se pudo restaurar una base de datos SQLite valida.")
    return plain_bytes


def _replace_file_atomically(target_path: Path, file_bytes):
    tmp_path = target_path.with_suffix(target_path.suffix + ".tmp")
    with tmp_path.open("wb") as f:
        f.write(file_bytes)
    os.replace(tmp_path, target_path)


def encrypt_database_file(password):
    if not DB_PATH.exists():
        return False
    raw_bytes = DB_PATH.read_bytes()
    if _is_db_file_encrypted(DB_PATH):
        return False
    if not raw_bytes.startswith(b"SQLite format 3\x00"):
        raise ValueError("El fichero de base de datos no es SQLite valido.")
    _replace_file_atomically(DB_PATH, _encrypt_db_payload(raw_bytes, password))
    return True


def decrypt_database_file(password):
    if not DB_PATH.exists():
        return False
    raw_bytes = DB_PATH.read_bytes()
    if not raw_bytes.startswith(DB_ENCRYPTION_MAGIC):
        return False
    plain_bytes = _decrypt_db_payload(raw_bytes, password)
    _replace_file_atomically(DB_PATH, plain_bytes)
    return True


def prepare_database_for_runtime(parent=None):
    global RUNTIME_DB_ENCRYPTION_ENABLED, RUNTIME_DB_ENCRYPTION_PASSWORD

    RUNTIME_DB_ENCRYPTION_ENABLED = _db_encryption_enabled_in_config()
    if not RUNTIME_DB_ENCRYPTION_ENABLED:
        RUNTIME_DB_ENCRYPTION_PASSWORD = None
        return True

    password, ok = QInputDialog.getText(
        parent,
        "Encriptacion de la base de datos",
        "Introduzca su contraseña para desbloquear la base de datos:",
        QLineEdit.EchoMode.Password,
    )
    if not ok:
        return False
    password = (password or "").strip()
    if not password:
        QMessageBox.warning(
            parent,
            "Clave obligatoria",
            "Debe introducir una contraseña valida para abrir PreCita.",
        )
        return False
    try:
        decrypt_database_file(password)
    except (ValueError, OSError) as exc:
        QMessageBox.critical(
            parent,
            "No se pudo abrir la base de datos",
            f"La contraseña no es valida o el fichero cifrado esta dañado.\n\nDetalle: {exc}",
        )
        return False

    RUNTIME_DB_ENCRYPTION_PASSWORD = password
    return True


def finalize_database_encryption_on_exit():
    if not RUNTIME_DB_ENCRYPTION_ENABLED:
        return
    if not RUNTIME_DB_ENCRYPTION_PASSWORD:
        return
    try:
        encrypt_database_file(RUNTIME_DB_ENCRYPTION_PASSWORD)
    except (ValueError, OSError):
        pass


def stylesheet_for_theme(theme):
    return PRECITA_QSS_DARK if theme == "dark" else PRECITA_QSS_LIGHT


def get_display_scale_percent():
    raw_scale = get_setting(
        "display_scale_percent", str(DEFAULT_DISPLAY_SCALE_PERCENT)
    )
    try:
        scale_percent = int(raw_scale)
    except (TypeError, ValueError):
        return DEFAULT_DISPLAY_SCALE_PERCENT
    if scale_percent in DISPLAY_SCALE_CHOICES_PERCENT:
        return scale_percent
    return DEFAULT_DISPLAY_SCALE_PERCENT


def _scaled_px_value(match, scale_ratio):
    value = float(match.group(1))
    scaled = value * scale_ratio
    if abs(scaled) < 0.0001:
        scaled = 0.0
    if abs(scaled - round(scaled)) < 0.001:
        formatted = str(int(round(scaled)))
    else:
        formatted = f"{scaled:.2f}".rstrip("0").rstrip(".")
    return f"{formatted}px"


def stylesheet_for_appearance(theme, scale_percent):
    base_stylesheet = stylesheet_for_theme(theme)
    scale_ratio = max(0.5, min(2.5, float(scale_percent) / 100.0))
    return re.sub(
        r"(-?\d+(?:\.\d+)?)px",
        lambda m: _scaled_px_value(m, scale_ratio),
        base_stylesheet,
    )


def build_app_font(scale_percent):
    scale_ratio = max(0.5, min(2.5, float(scale_percent) / 100.0))
    base_size = max(8, int(round(10 * scale_ratio)))
    font = QFont("Google Sans", base_size)
    if font.family() != "Google Sans":
        font = QFont("Segoe UI", base_size)
    if font.family() not in {"Google Sans", "Segoe UI"}:
        font = QFont()
        font.setPointSize(base_size)
    return font


def apply_app_appearance(app, theme, scale_percent):
    if not app:
        return
    app.setStyleSheet(stylesheet_for_appearance(theme, scale_percent))
    app.setFont(build_app_font(scale_percent))


def get_startup_command():
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    return f'"{sys.executable}" "{Path(__file__).resolve()}"'


def windows_startup_is_enabled():
    if sys.platform != "win32":
        return False
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_READ,
        )
        try:
            winreg.QueryValueEx(key, "PreCita")
            return True
        except OSError:
            return False
        finally:
            key.Close()
    except OSError:
        return False


def windows_startup_set(enabled):
    if sys.platform != "win32":
        return False, "Solo está disponible en Windows."
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    name = "PreCita"
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            key_path,
            0,
            winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE,
        )
        try:
            if enabled:
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, get_startup_command())
            else:
                try:
                    winreg.DeleteValue(key, name)
                except FileNotFoundError:
                    pass
        finally:
            key.Close()
        return True, ""
    except OSError as e:
        return False, str(e)


def sync_windows_startup_with_settings():
    """Sincroniza el inicio con Windows según la configuración guardada."""
    if sys.platform != "win32":
        return True, ""
    enabled = get_setting("windows_startup", DEFAULT_WINDOWS_STARTUP) == "1"
    return windows_startup_set(enabled)


def revoke_and_remove_google_credentials() -> None:
    """Revoca el token en Google (si es posible) y borra token.json local."""
    if not CREDENTIALS_PATH.exists():
        return
    try:
        creds = Credentials.from_authorized_user_file(str(CREDENTIALS_PATH), SCOPES)
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                pass
        token = creds.token or creds.refresh_token
        if token:
            data = urllib.parse.urlencode({"token": token}).encode()
            req = urllib.request.Request(
                "https://oauth2.googleapis.com/revoke",
                data=data,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                method="POST",
            )
            try:
                urllib.request.urlopen(req, timeout=15)
            except (urllib.error.URLError, urllib.error.HTTPError):
                pass
    except Exception:
        pass
    try:
        CREDENTIALS_PATH.unlink(missing_ok=True)
    except OSError:
        pass

# ============================================================================
# AUTENTICACIÓN CON GOOGLE (Calendar + Gmail)
# ============================================================================

class GoogleOAuthDialog(QDialog):
    """Ventana embebida para iniciar sesión OAuth de Google dentro de PreCita."""

    def __init__(self, flow: InstalledAppFlow, parent=None):
        super().__init__(parent)
        self.flow = flow
        self.authorization_response_url = None
        self.oauth_error = None
        self.opened_in_external_browser = False

        self.setWindowTitle("Iniciar sesión con Google")
        self.resize(760, 620)
        self.setMinimumSize(680, 520)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(4)

        transparency_hint = QLabel(
            "PreCita muestra aquí la URL real cargada en esta ventana."
        )
        transparency_hint.setObjectName("oauthTransparencyHint")
        transparency_hint.setWordWrap(False)
        transparency_hint.setFixedHeight(14)
        transparency_hint.setStyleSheet(
            "font-size: 9px; color: #5a6b80; margin: 0; padding: 0;"
        )
        layout.addWidget(transparency_hint)

        url_row = QHBoxLayout()
        url_row.setContentsMargins(0, 0, 0, 0)
        url_row.setSpacing(6)

        self.current_url_display = QLineEdit(self)
        self.current_url_display.setReadOnly(True)
        self.current_url_display.setPlaceholderText("URL actual")
        self.current_url_display.setToolTip(
            "Revise esta URL para confirmar que pertenece al dominio esperado."
        )
        url_row.addWidget(self.current_url_display, 1)

        self.open_browser_btn = QPushButton("Abrir en el navegador")
        self.open_browser_btn.clicked.connect(self._open_in_system_browser)
        url_row.addWidget(self.open_browser_btn)

        layout.addLayout(url_row)

        self.web_view = QWebEngineView(self)
        self.web_view.urlChanged.connect(self._on_url_changed)
        layout.addWidget(self.web_view)

        auth_url, _ = self.flow.authorization_url(
            access_type="offline",
            include_granted_scopes="true",
            prompt="consent",
            hl="es",
        )
        self.current_url_display.setText(auth_url)
        self.web_view.setUrl(QUrl(auth_url))

    def _open_in_system_browser(self):
        url = self.current_url_display.text().strip()
        if not url:
            return
        warning_box = QMessageBox(self)
        warning_box.setIcon(QMessageBox.Icon.Warning)
        warning_box.setWindowTitle("Aviso importante de sincronización")
        warning_box.setText(
            "Se abrirá Google en su navegador predeterminado.\n\n"
            "No cierre esa pestaña hasta finalizar la autorización.\n"
            "Si la cierra antes de completar la sincronización, PreCita puede quedarse "
            "bloqueado y tendrá que cerrarlo y abrirlo de nuevo."
        )
        open_button = warning_box.addButton(
            "Abrir en el navegador", QMessageBox.ButtonRole.AcceptRole
        )
        warning_box.addButton("Cancelar", QMessageBox.ButtonRole.RejectRole)
        warning_box.setDefaultButton(open_button)
        warning_box.exec()
        if warning_box.clickedButton() is not open_button:
            # Incluye pulsar "Cancelar" o cerrar con la X.
            return
        QDesktopServices.openUrl(QUrl(url))
        self.opened_in_external_browser = True
        self.reject()

    def _on_url_changed(self, qurl: QUrl):
        url = qurl.toString()
        self.current_url_display.setText(url)
        if not (
            url.startswith(f"http://localhost:{OAUTH_LOOPBACK_PORT}/")
            or url.startswith(f"http://127.0.0.1:{OAUTH_LOOPBACK_PORT}/")
        ):
            return

        parsed = urllib.parse.urlparse(url)
        query = urllib.parse.parse_qs(parsed.query)

        if "error" in query:
            self.oauth_error = query.get("error", ["Error desconocido"])[0]
            self.accept()
            return

        if "code" in query:
            self.authorization_response_url = url
            self.accept()


def _run_embedded_google_oauth(flow: InstalledAppFlow, parent=None) -> Credentials:
    """Ejecuta el login OAuth en una ventana embebida y retorna credenciales."""
    dialog = GoogleOAuthDialog(flow, parent)
    result = dialog.exec()
    if dialog.opened_in_external_browser:
        # Continúa el flujo OAuth en el navegador predeterminado ya abierto.
        # Evita bloqueo indefinido cuando el usuario cierra la pestaña sin autorizar.
        try:
            return _run_external_browser_oauth(
                flow, timeout_seconds=OAUTH_EXTERNAL_TIMEOUT_SECONDS
            )
        except TimeoutError as e:
            raise RuntimeError(
                "No se completó la autorización en el navegador (tiempo agotado). "
                "Si cerró la pestaña de Google, vuelva a pulsar Sincronizar."
            ) from e
        except Exception as e:
            raise RuntimeError(
                "No se completó el inicio de sesión con Google desde el navegador. "
                "Si cerró la pestaña o no autorizó, inténtelo de nuevo."
            ) from e
    if result != QDialog.DialogCode.Accepted:
        raise RuntimeError("Inicio de sesión de Google cancelado por el usuario.")
    if dialog.oauth_error:
        raise RuntimeError(f"Google rechazó el acceso: {dialog.oauth_error}")
    if not dialog.authorization_response_url:
        raise RuntimeError("No se recibió el código de autorización de Google.")

    flow.fetch_token(authorization_response=dialog.authorization_response_url)
    return flow.credentials


def _run_external_browser_oauth(flow: InstalledAppFlow, timeout_seconds: int = 90) -> Credentials:
    """Ejecuta OAuth en navegador externo sin bloquear indefinidamente la UI."""
    _ensure_loopback_port_available(OAUTH_LOOPBACK_PORT)
    kwargs = {
        "port": OAUTH_LOOPBACK_PORT,
        "open_browser": False,
        "authorization_prompt_message": "",
        "success_message": (
            "PreCita: sincronización con Google completada. "
            "Puede cerrar esta pestaña."
        ),
    }
    run_local_server_params = inspect.signature(flow.run_local_server).parameters
    if "timeout_seconds" in run_local_server_params:
        kwargs["timeout_seconds"] = timeout_seconds
        return flow.run_local_server(**kwargs)

    # Compatibilidad con versiones antiguas: correr en hilo aparte y aplicar timeout manual.
    result_queue = queue.Queue(maxsize=1)

    def _target():
        try:
            creds = flow.run_local_server(**kwargs)
            result_queue.put(("ok", creds))
        except Exception as err:  # pragma: no cover - ruta de error en tiempo real
            result_queue.put(("error", err))

    worker = threading.Thread(target=_target, daemon=True)
    worker.start()

    deadline = datetime.now() + timedelta(seconds=timeout_seconds)
    while datetime.now() < deadline:
        QApplication.processEvents()
        try:
            status, payload = result_queue.get(timeout=0.1)
        except queue.Empty:
            continue
        if status == "error":
            raise payload
        return payload

    raise TimeoutError("OAuth timeout reached")


def _ensure_loopback_port_available(port: int) -> None:
    """Verifica que el puerto local requerido para OAuth esté libre."""
    probe = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        probe.settimeout(0.5)
        probe.bind(("127.0.0.1", port))
    except OSError as e:
        raise RuntimeError(
            f"El puerto local {port} ya está en uso por otro proceso. "
            "Cierre ese proceso o cambie su configuración para poder autenticar con Google."
        ) from e
    finally:
        probe.close()


def load_google_client_config():
    """Carga y descifra la configuración de Google usando una clave de entorno."""
    try:
        if CLIENT_SECRETS.exists():
            f = Fernet(PRECITA_MASTER_KEY)            
            with open(CLIENT_SECRETS, 'rb') as bf:
                encrypted_content = bf.read()
            decrypted_content = f.decrypt(encrypted_content)
            return json.loads(decrypted_content.decode('utf-8'))
        else:
            print(f"Error: El archivo {CLIENT_SECRETS} no existe.")
            return None

    except Exception as e:
        print(f"Error crítico en la carga de configuración.")
        return None


def get_google_service(scope_type='calendar', embedded_oauth=False, parent=None):
    """Obtiene el servicio de Google (Calendar o Gmail) con autenticación OAuth2."""
    creds = None
    
    if scope_type == 'calendar':
        scopes = SCOPES_CALENDAR
        service_name = 'calendar'
        service_version = 'v3'
    else:
        scopes = SCOPES_GMAIL
        service_name = 'gmail'
        service_version = 'v1'
    
    # Si existe token guardado, usarlo
    if CREDENTIALS_PATH.exists():
        creds = Credentials.from_authorized_user_file(
            str(CREDENTIALS_PATH), SCOPES
        )
    
    # Si las credenciales expiaron, refrescar
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    
    # Si no hay credenciales válidas, hacer login OAuth
    if not creds or not creds.valid:
        client_config = load_google_client_config()
        flow = InstalledAppFlow.from_client_config(client_config, scopes)
        flow.redirect_uri = f"http://localhost:{OAUTH_LOOPBACK_PORT}/"

        if embedded_oauth:
            creds = _run_embedded_google_oauth(flow, parent=parent)
        else:
            _ensure_loopback_port_available(OAUTH_LOOPBACK_PORT)
            creds = flow.run_local_server(port=OAUTH_LOOPBACK_PORT)
        
        # Guardar credenciales para futuras ejecuciones
        with open(CREDENTIALS_PATH, 'w', encoding='utf-8') as token:
            token.write(creds.to_json())
    
    return build(service_name, service_version, credentials=creds)


def is_google_session_synced() -> bool:
    """Indica si hay una sesión de Google válida guardada localmente."""
    if not CREDENTIALS_PATH.exists():
        return False
    try:
        creds = Credentials.from_authorized_user_file(str(CREDENTIALS_PATH), SCOPES)
        if creds.valid:
            return True
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                if creds.valid:
                    with open(CREDENTIALS_PATH, "w", encoding="utf-8") as token:
                        token.write(creds.to_json())
                    return True
            except Exception:
                return False
    except Exception:
        return False
    return False

# ============================================================================
# FUNCIONES DE EMAIL (Gmail API)
# ============================================================================

def format_email_template(template, variables):
    """Reemplaza variables {clave} en una plantilla de email."""
    rendered = template
    for key, value in variables.items():
        rendered = rendered.replace(f"{{{key}}}", str(value or ""))
    return rendered


def _build_email_template_variables(
    contact_first_name,
    contact_last_name,
    contact_email,
    contact_phone,
    event_date,
):
    """Construye variables disponibles para plantilla de recordatorio."""
    nombre_citado = (contact_first_name or "").strip()
    apellidos_citado = (contact_last_name or "").strip()
    correo_citado = (contact_email or "").strip()
    tlf_citado = (contact_phone or "").strip()

    hora_cita = event_date
    fecha_cita = ""
    dia_semana = ""
    try:
        apt_datetime = datetime.fromisoformat(event_date.replace("Z", "+00:00"))
        hora_cita = apt_datetime.strftime("%H:%M")
        fecha_cita = apt_datetime.strftime("%d/%m/%Y")
        dia_semana = (
            "lunes",
            "martes",
            "miercoles",
            "jueves",
            "viernes",
            "sabado",
            "domingo",
        )[apt_datetime.weekday()]
    except Exception:
        pass

    return {
        "nombre_citado": nombre_citado,
        "apellidos_citado": apellidos_citado,
        "correo_citado": correo_citado,
        "tlf_citado": tlf_citado,
        "hora_cita": hora_cita,
        "fecha_cita": fecha_cita,
        "dia_semana": dia_semana,
    }


def _load_email_template_settings():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        '''
        SELECT key, value FROM settings
        WHERE key IN (
            "email_template",
            "email_subject_template",
            "email_template_format",
            "email_template_attachments"
        )
        '''
    )
    rows = dict(cursor.fetchall())
    conn.close()
    return rows


def _blocked_extension_of(name):
    return Path(name).suffix.strip().lower()


def _zip_contains_blocked_extension(zip_path):
    try:
        with zipfile.ZipFile(zip_path, "r") as archive:
            for member in archive.infolist():
                if member.is_dir():
                    continue
                if _blocked_extension_of(member.filename) in BLOCKED_ATTACHMENT_EXTENSIONS:
                    return True, member.filename
    except zipfile.BadZipFile:
        return True, "ZIP corrupto"
    return False, ""


def _validate_attachment_security(file_path):
    path = Path(file_path).expanduser().resolve()
    if not path.is_file():
        return False, f"No se encuentra el archivo: {path}"

    ext = _blocked_extension_of(path.name)
    if ext in BLOCKED_ATTACHMENT_EXTENSIONS:
        return False, f"Tipo de archivo bloqueado por seguridad: {ext}\n\nRevisa 'Ayuda' para más información."

    if ext in ARCHIVE_EXTENSIONS:
        if ext == ".zip":
            has_blocked, blocked_member = _zip_contains_blocked_extension(path)
            if has_blocked:
                return False, f"ZIP bloqueado por contener: {blocked_member}\n\nRevisa 'Ayuda' para más información."
        else:
            return False, (
                f"Archivo comprimido no permitido para seguridad avanzada ({ext}). "
                "Use ZIP si necesita adjuntar comprimidos.\n\nRevisa 'Ayuda' para más información."
            )

    return True, ""


def _parse_template_attachments(raw_value):
    """Normaliza lista de adjuntos almacenada en settings."""
    if not raw_value:
        return []
    try:
        parsed = json.loads(raw_value)
    except Exception:
        return []
    if not isinstance(parsed, list):
        return []
    normalized = []
    for item in parsed:
        if isinstance(item, str) and item.strip():
            old_path = str(Path(item.strip()).expanduser())
            normalized.append(
                {
                    "path": old_path,
                    "name": Path(old_path).name,
                }
            )
        elif isinstance(item, dict):
            stored_path = str(item.get("path", "")).strip()
            file_name = str(item.get("name", "")).strip() or Path(stored_path).name
            if stored_path:
                is_inline = bool(item.get("inline", False))
                cid_value = str(item.get("cid", "")).strip() if is_inline else ""
                normalized.append(
                    {
                        "path": str(Path(stored_path).expanduser()),
                        "name": file_name,
                        "inline": is_inline,
                        "cid": cid_value,
                    }
                )
    return normalized


def _template_payload_size_bytes(body_html, attachment_items):
    size = len(body_html.encode("utf-8"))
    for item in attachment_items:
        stored_path = Path(item.get("path", "")).expanduser()
        if stored_path.is_file():
            size += stored_path.stat().st_size
    return size

def send_reminder_email_gmail(appointment_id, recipient_email, variables):
    """Envía un email de recordatorio usando Gmail API."""
    try:
        gmail_service = get_google_service('gmail')
        
        template_settings = _load_email_template_settings()
        template = template_settings.get("email_template", "")
        subject_template = template_settings.get(
            "email_subject_template", "Recordatorio de cita - {nombre_citado} {apellidos_citado}"
        )
        body_format = template_settings.get("email_template_format", "plain")
        attachment_items = _parse_template_attachments(
            template_settings.get("email_template_attachments", "[]")
        )

        # Formatear email
        body = format_email_template(template, variables)
        subject = format_email_template(subject_template, variables)
        payload_bytes = _template_payload_size_bytes(body, attachment_items)
        if payload_bytes > MAX_TEMPLATE_PAYLOAD_BYTES:
            return False, (
                "Adjuntos + cuerpo superan el límite de 16 MB. "
                "Reduzca adjuntos o contenido de plantilla."
            )

        # Crear mensaje. Si hay imágenes integradas, usamos multipart/related
        # para que Gmail las trate como inline (CID).
        inline_items = [item for item in attachment_items if item.get("inline")]
        regular_attachments = [item for item in attachment_items if not item.get("inline")]

        message = MIMEMultipart("mixed")
        message['to'] = recipient_email
        message['subject'] = subject

        body_subtype = "html" if body_format == "html" else "plain"
        if inline_items:
            related_part = MIMEMultipart("related")
            related_part.attach(MIMEText(body, "html"))
            for item in inline_items:
                path = Path(item.get("path", "")).expanduser()
                cid_value = str(item.get("cid", "")).strip()
                if not path.is_file() or not cid_value:
                    continue
                is_valid, reason = _validate_attachment_security(path)
                if not is_valid:
                    return False, f"Imagen integrada bloqueada: {reason}"
                with open(path, "rb") as file_obj:
                    image_data = file_obj.read()
                mime_type, _ = mimetypes.guess_type(str(path))
                subtype = None
                if mime_type and mime_type.startswith("image/"):
                    subtype = mime_type.split("/", 1)[1]
                image_part = MIMEImage(image_data, _subtype=subtype)
                display_name = item.get("name", path.name)
                image_part.add_header("Content-ID", f"<{cid_value}>")
                image_part.add_header(
                    "Content-Disposition",
                    f'inline; filename="{display_name}"'
                )
                related_part.attach(image_part)
            message.attach(related_part)
        else:
            message.attach(MIMEText(body, body_subtype))

        for item in regular_attachments:
            path = Path(item.get("path", "")).expanduser()
            if not path.is_file():
                continue
            is_valid, reason = _validate_attachment_security(path)
            if not is_valid:
                return False, f"Adjunto bloqueado: {reason}"
            mime_type, _ = mimetypes.guess_type(str(path))
            if mime_type and "/" in mime_type:
                maintype, subtype = mime_type.split("/", 1)
            else:
                maintype, subtype = "application", "octet-stream"
            with open(path, "rb") as file_obj:
                part = MIMEBase(maintype, subtype)
                part.set_payload(file_obj.read())
            encoders.encode_base64(part)
            display_name = item.get("name", path.name)
            part.add_header("Content-Disposition", f'attachment; filename="{display_name}"')
            message.attach(part)
        
        # Codificar en base64
        raw = base64.urlsafe_b64encode(message.as_bytes())
        raw = raw.decode()
        
        # Enviar a través de Gmail API
        send_message = {
            'raw': raw
        }
        
        gmail_service.users().messages().send(userId='me', body=send_message).execute()
        
        # Actualizar BD
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE appointments 
            SET reminder_sent = 1, email_sent_at = ?
            WHERE id = ?
        ''', (datetime.now().isoformat(), appointment_id))
        conn.commit()
        conn.close()
        
        return True, "Email enviado correctamente"
    
    except Exception as e:
        return False, f"Error al enviar email: {str(e)}"

# ============================================================================
# TRABAJADOR DE SINCRONIZACIÓN
# ============================================================================

class SyncWorker(QObject):
    """Worker para sincronizar citas sin bloquear la GUI."""
    finished = pyqtSignal()
    error = pyqtSignal(str)
    appointments_found = pyqtSignal(list)
    
    def run(self):
        try:
            service = get_google_service('calendar')
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            calendar_id, calendar_name = get_selected_google_calendar()
            
            # Obtener citas de los próximos días configurados
            now = datetime.now(pytz.UTC)
            sync_days = get_calendar_sync_days()
            sync_until = now + timedelta(days=sync_days)
            
            events_result = service.events().list(
                calendarId=calendar_id,
                timeMin=now.isoformat(),
                timeMax=sync_until.isoformat(),
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            
            events = events_result.get('items', [])
            appointments = []
            
            for event in events:
                event_id = event['id']
                event_title = event.get('summary', 'Sin título')
                event_start = event['start'].get('dateTime', event['start'].get('date'))
                event_description = event.get('description', '')
                
                # Verificar si ya existe en BD
                cursor.execute(
                    'SELECT id FROM appointments WHERE calendar_event_id = ? AND source_calendar_id = ?',
                    (event_id, calendar_id)
                )
                
                if not cursor.fetchone():
                    # Insertar nueva cita
                    cursor.execute('''
                        INSERT INTO appointments 
                        (calendar_event_id, source_calendar_id, event_title, event_date, event_description)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (event_id, calendar_id, event_title, event_start, event_description))
                    conn.commit()
                    appointments.append({
                        'title': event_title,
                        'date': event_start,
                        'id': event_id
                    })
            
            self.appointments_found.emit(appointments)
            conn.close()
            
        except Exception as e:
            details = str(e)
            if "Not Found" in details or "404" in details:
                self.error.emit(
                    f"Error en sincronización: el calendario configurado "
                    f"({calendar_name}) no está disponible. Revise la selección "
                    "de calendario en Configuración."
                )
            else:
                self.error.emit(f"Error en sincronización: {details}")
        finally:
            self.finished.emit()

# ============================================================================
# TRABAJADOR DE RECORDATORIOS
# ============================================================================

class ReminderWorker(QObject):
    """Worker para enviar recordatorios de citas automáticamente."""
    finished = pyqtSignal()
    error = pyqtSignal(str)
    reminder_sent = pyqtSignal(str)  # Emite mensaje de cita procesada
    auto_reminder_email_sent = pyqtSignal(str, str)  # contacto, hora cita

    def __init__(self, auto_mode=False):
        super().__init__()
        self.auto_mode = auto_mode
    
    def run(self):
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            
            # Obtener citas para mañana que no han recibido recordatorio
            tomorrow = datetime.now() + timedelta(days=1)
            tomorrow_start = tomorrow.replace(hour=0, minute=0, second=0, microsecond=0)
            tomorrow_end = tomorrow.replace(hour=23, minute=59, second=59, microsecond=0)
            
            cursor.execute('''
                SELECT a.id, a.event_title, a.event_date, a.contact_id,
                       c.first_name, c.last_name, c.email, c.phone
                FROM appointments a
                LEFT JOIN contacts c ON a.contact_id = c.id
                WHERE a.reminder_sent = 0
                AND datetime(a.event_date) >= datetime(?)
                AND datetime(a.event_date) <= datetime(?)
            ''', (tomorrow_start.isoformat(), tomorrow_end.isoformat()))
            
            pending_appointments = cursor.fetchall()
            conn.close()
            
            if not pending_appointments:
                self.reminder_sent.emit("ℹ️ No hay citas para mañana")
                self.finished.emit()
                return
            
            for apt in pending_appointments:
                (
                    apt_id,
                    title,
                    event_date,
                    contact_id,
                    c_first,
                    c_last,
                    contact_email,
                    contact_phone,
                ) = apt
                contact_display_name = contact_full_name(c_first, c_last) or "Paciente"
                
                # Si no hay contacto asociado, saltarlo
                if not contact_email:
                    self.reminder_sent.emit(f"⏭️ {title}: Sin email asociado")
                    continue
                
                template_vars = _build_email_template_variables(
                    c_first,
                    c_last,
                    contact_email,
                    contact_phone,
                    event_date,
                )
                # Enviar email
                success, message = send_reminder_email_gmail(
                    apt_id,
                    contact_email,
                    template_vars,
                )
                
                if success:
                    self.reminder_sent.emit(f"✓ Email enviado a {contact_display_name}")
                    if self.auto_mode:
                        self.auto_reminder_email_sent.emit(
                            contact_display_name, template_vars.get("hora_cita", event_date)
                        )
                else:
                    self.reminder_sent.emit(f"✗ Error: {message}")
            
        except Exception as e:
            self.error.emit(f"Error en recordatorios: {str(e)}")
        finally:
            self.finished.emit()

# ============================================================================
# EDITOR RICH TEXT CON RESIZE VISUAL DE IMAGEN
# ============================================================================

class ImageEditableTextEdit(QTextEdit):
    """QTextEdit con marco y tiradores para redimensionar imágenes seleccionadas."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self._selected_image_src = ""
        self._selected_image_pos = -1
        self._selected_image_rect = QRect()
        self._dragging_handle = ""
        self._drag_start_pos = QPoint()
        self._drag_start_width = 0
        self._resize_callback = None

    def set_image_resize_callback(self, callback):
        self._resize_callback = callback

    def _image_cursor_from_position(self, cursor):
        """Devuelve un cursor posicionado sobre una imagen si existe en la posición o justo antes."""
        if cursor.charFormat().isImageFormat():
            return QTextCursor(cursor)
        if cursor.position() > 0:
            previous_cursor = QTextCursor(cursor)
            previous_cursor.setPosition(cursor.position() - 1)
            if previous_cursor.charFormat().isImageFormat():
                return previous_cursor
        return None

    def _compute_image_rect(self, cursor):
        """Calcula un rectángulo visual estable para la imagen seleccionada."""
        char_format = cursor.charFormat()
        if not char_format.isImageFormat():
            return QRect()
        image_format = char_format.toImageFormat()
        caret_rect = self.cursorRect(cursor).adjusted(0, 0, -1, -1)

        width = int(round(image_format.width() if image_format.width() > 0 else 0))
        height = int(round(image_format.height() if image_format.height() > 0 else 0))

        if width <= 0 or height <= 0:
            image_resource = self.document().resource(
                QTextDocument.ResourceType.ImageResource,
                QUrl(str(image_format.name() or "")),
            )
            native_width = 0
            native_height = 0
            if isinstance(image_resource, QImage):
                native_width = image_resource.width()
                native_height = image_resource.height()
            elif hasattr(image_resource, "size"):
                size = image_resource.size()
                native_width = int(size.width())
                native_height = int(size.height())
            if width <= 0:
                width = native_width
            if height <= 0:
                height = native_height

        width = max(30, min(width or 1, int(self.viewport().width() * 0.95)))
        if height <= 0:
            height = max(24, int(width * 0.6))

        return QRect(caret_rect.left(), caret_rect.top(), width, height)

    def select_image_by_src(self, src):
        """Re-selecciona una imagen por src tras reconstrucciones de HTML."""
        if not src:
            self.clear_image_selection()
            return False
        cursor = QTextCursor(self.document())
        while not cursor.atEnd():
            char_format = cursor.charFormat()
            if char_format.isImageFormat():
                image_format = char_format.toImageFormat()
                current_src = str(image_format.name() or "").strip()
                if current_src == src:
                    self._set_selected_image_from_cursor(cursor)
                    return True
            if not cursor.movePosition(QTextCursor.MoveOperation.NextCharacter):
                break
        self.clear_image_selection()
        return False

    def clear_image_selection(self):
        self._selected_image_src = ""
        self._selected_image_pos = -1
        self._selected_image_rect = QRect()
        self._dragging_handle = ""
        self.viewport().update()

    def _set_selected_image_from_cursor(self, cursor):
        image_cursor = self._image_cursor_from_position(cursor)
        if image_cursor is None:
            self.clear_image_selection()
            return False
        char_format = image_cursor.charFormat()
        if not char_format.isImageFormat():
            self.clear_image_selection()
            return False
        image_format = char_format.toImageFormat()
        src = str(image_format.name() or "").strip()
        if not (src.startswith("data:image/") or src.startswith("cid:")):
            self.clear_image_selection()
            return False
        self._selected_image_src = src
        self._selected_image_pos = image_cursor.position()
        self._selected_image_rect = self._compute_image_rect(image_cursor)
        selected_cursor = QTextCursor(image_cursor)
        selected_cursor.setPosition(self._selected_image_pos)
        if self._selected_image_pos < self.document().characterCount() - 1:
            selected_cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
        self.setTextCursor(selected_cursor)
        self.viewport().update()
        return True

    def _handle_points(self):
        rect = self._selected_image_rect
        if rect.isNull():
            return {}
        return {
            "top_left": rect.topLeft(),
            "top_center": QPoint(rect.center().x(), rect.top()),
            "top_right": rect.topRight(),
            "middle_left": QPoint(rect.left(), rect.center().y()),
            "middle_right": QPoint(rect.right(), rect.center().y()),
            "bottom_left": rect.bottomLeft(),
            "bottom_right": rect.bottomRight(),
            "bottom_center": QPoint(rect.center().x(), rect.bottom()),
        }

    def _handle_rect(self, point):
        size = 10
        return QRect(point.x() - size // 2, point.y() - size // 2, size, size)

    def _handle_at(self, pos):
        if self._selected_image_rect.isNull():
            return ""
        for handle_name, point in self._handle_points().items():
            if self._handle_rect(point).contains(pos):
                return handle_name
        return ""

    def _refresh_selected_image_rect(self):
        if self._selected_image_pos < 0:
            return
        cursor = self.textCursor()
        cursor.setPosition(self._selected_image_pos)
        if not cursor.charFormat().isImageFormat():
            self.clear_image_selection()
            return
        self._selected_image_rect = self._compute_image_rect(cursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            handle_name = self._handle_at(event.pos())
            if handle_name and self._selected_image_src:
                self._dragging_handle = handle_name
                self._drag_start_pos = event.pos()
                self._refresh_selected_image_rect()
                self._drag_start_width = max(1, self._selected_image_rect.width())
                event.accept()
                return
            cursor = self.cursorForPosition(event.pos())
            if self._set_selected_image_from_cursor(cursor):
                event.accept()
                return
            self.clear_image_selection()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._dragging_handle and self._selected_image_src and self._resize_callback:
            delta_x = event.pos().x() - self._drag_start_pos.x()
            delta_y = event.pos().y() - self._drag_start_pos.y()
            if self._dragging_handle in ("top_left", "middle_left", "bottom_left"):
                delta_x = -delta_x
            if self._dragging_handle in ("top_center",):
                primary_delta = -delta_y
            elif self._dragging_handle in ("bottom_center",):
                primary_delta = delta_y
            elif self._dragging_handle in ("middle_left", "middle_right"):
                primary_delta = delta_x
            else:
                primary_delta = delta_x if abs(delta_x) >= abs(delta_y) else delta_y
            scale_factor = 1.0 + (primary_delta / max(1, self._drag_start_width))
            scale_factor = max(0.1, min(4.0, scale_factor))
            self._resize_callback(self._selected_image_src, scale_factor)
            self._refresh_selected_image_rect()
            self.viewport().update()
            event.accept()
            return

        handle_name = self._handle_at(event.pos())
        if handle_name in ("middle_left", "middle_right"):
            self.viewport().setCursor(Qt.CursorShape.SizeHorCursor)
        elif handle_name in ("top_center", "bottom_center"):
            self.viewport().setCursor(Qt.CursorShape.SizeVerCursor)
        elif handle_name in ("top_left", "bottom_right"):
            self.viewport().setCursor(Qt.CursorShape.SizeFDiagCursor)
        elif handle_name in ("top_right", "bottom_left"):
            self.viewport().setCursor(Qt.CursorShape.SizeBDiagCursor)
        elif self._selected_image_rect.contains(event.pos()):
            self.viewport().setCursor(Qt.CursorShape.SizeAllCursor)
        else:
            self.viewport().unsetCursor()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self._dragging_handle:
            self._dragging_handle = ""
            self.viewport().unsetCursor()
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._selected_image_rect.isNull():
            return
        painter = QPainter(self.viewport())
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor("#2e86de"), 1))
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRect(self._selected_image_rect)
        painter.setPen(QPen(QColor("#2e86de"), 1))
        painter.setBrush(QBrush(QColor("#ffffff")))
        for point in self._handle_points().values():
            painter.drawRect(self._handle_rect(point))
        painter.end()


# ============================================================================
# DIÁLOGO DE EDICIÓN DE PLANTILLA
# ============================================================================

class TemplateEditorDialog(QDialog):
    """Diálogo para editar la plantilla de email."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Personalizar plantilla — PreCita — {get_version()}")
        self.setMinimumSize(980, 620)
        self.resize(1120, 720)
        
        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(12)
        
        card = QFrame()
        card.setObjectName("panelCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 16, 16, 16)
        card_layout.setSpacing(14)

        content_layout = QHBoxLayout()
        content_layout.setSpacing(16)
        card_layout.addLayout(content_layout)

        left_panel = QGroupBox("Variables disponibles")
        left_panel.setMinimumWidth(320)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(12, 10, 12, 12)
        left_layout.setSpacing(10)

        vars_help = QLabel()
        vars_help.setTextFormat(Qt.TextFormat.RichText)
        vars_help.setWordWrap(True)
        vars_help.setText(
            "<p style='margin-top:0; margin-bottom:0; line-height:1.45;'>"
            "Aquí define el formato del correo que enviará PreCita. "
            "Use texto libre y los parámetros con llaves. "
            "No tiene por qué usar todas las variables. <br><br>"
            "<b>{nombre_citado}</b> — Nombre del contacto vinculado a la cita.<br><br>"
            "<b>{apellidos_citado}</b> — Apellido(s) del contacto vinculado a la cita.<br><br>"
            "<b>{correo_citado}</b> — Correo electrónico del contacto vinculado a la cita.<br><br>"
            "<b>{tlf_citado}</b> — Teléfono del contacto vinculado a la cita. Si usa esta variable, asegúrese de tener guardados todos los teléfonos de todos los contactos.<br><br>"
            "<b>{hora_cita}</b> — Hora de la cita en formato 24&nbsp;h (HH:MM).<br><br>"
            "<b>{fecha_cita}</b> — Fecha de la cita en formato DD/MM/YYYY.<br><br>"
            "<b>{dia_semana}</b> — Día de la semana de la cita (lunes, martes, etc.)."
            "</p>"
        )
        left_layout.addWidget(vars_help)
        left_layout.addStretch()
        content_layout.addWidget(left_panel, 0)

        right_panel = QFrame()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)
        content_layout.addWidget(right_panel, 1)

        subject_title = QLabel("Asunto")
        subject_title.setObjectName("sectionTitle")
        right_layout.addWidget(subject_title)

        self.subject_input = QLineEdit()
        self.subject_input.setPlaceholderText(
            "Ej.: Recordatorio de cita - {nombre_citado} {apellidos_citado}"
        )
        right_layout.addWidget(self.subject_input)

        body_title = QLabel("Cuerpo del mensaje")
        body_title.setObjectName("sectionTitle")
        right_layout.addWidget(body_title)

        toolbar = QHBoxLayout()
        toolbar.setSpacing(6)
        bold_btn = QPushButton("N")
        bold_btn.setToolTip("Negrita")
        bold_btn.setObjectName("formatToggleButton")
        bold_btn.setCheckable(True)
        italic_btn = QPushButton("I")
        italic_btn.setToolTip("Cursiva")
        italic_btn.setObjectName("formatToggleButton")
        italic_btn.setCheckable(True)
        underline_btn = QPushButton("S")
        underline_btn.setToolTip("Subrayado")
        underline_btn.setObjectName("formatToggleButton")
        underline_btn.setCheckable(True)
        bold_font = bold_btn.font()
        bold_font.setBold(True)
        bold_btn.setFont(bold_font)
        bold_btn.setStyleSheet("font-weight: 700;")
        italic_font = italic_btn.font()
        italic_font.setItalic(True)
        italic_btn.setFont(italic_font)
        underline_font = underline_btn.font()
        underline_font.setUnderline(True)
        underline_btn.setFont(underline_font)
        bullet_btn = QPushButton("•")
        bullet_btn.setToolTip("Lista con viñetas")
        bullet_btn.setObjectName("formatToggleButton")
        bullet_btn.setCheckable(True)
        clear_btn = QPushButton("Borrar formato")
        clear_btn.setToolTip("Conservar texto y limpiar estilos")
        toolbar.addWidget(bold_btn)
        toolbar.addWidget(italic_btn)
        toolbar.addWidget(underline_btn)
        toolbar.addWidget(bullet_btn)
        toolbar.addWidget(clear_btn)
        toolbar.addStretch()
        attach_btn = QPushButton("Adjuntar archivos")
        attach_btn.setToolTip("Adjuntar archivos a todos los recordatorios")
        inline_image_btn = QPushButton("Insertar imagen")
        inline_image_btn.setObjectName("inlineImageButton")
        inline_image_btn.setProperty("softDisabled", True)
        inline_image_btn.setCursor(Qt.CursorShape.ForbiddenCursor)
        inline_image_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        inline_image_btn.setToolTip("Esta funcionalidad será añadida en versiones futuras")
        toolbar.addWidget(attach_btn)
        toolbar.addWidget(inline_image_btn)
        right_layout.addLayout(toolbar)

        self.template_text = ImageEditableTextEdit()
        self.template_text.setMinimumHeight(220)
        self.template_text.set_image_resize_callback(self._resize_selected_image_by_factor)
        self.template_text.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.template_text.customContextMenuRequested.connect(self._show_editor_context_menu)
        template_settings = _load_email_template_settings()
        self.subject_input.setText(
            template_settings.get(
                "email_subject_template", "Recordatorio de cita - {nombre_citado} {apellidos_citado}"
            )
        )
        saved_attachments = _parse_template_attachments(
            template_settings.get("email_template_attachments", "[]")
        )
        self.attachments = []
        for attachment_item in saved_attachments:
            self._add_attachment_item(attachment_item)

        template_format = template_settings.get("email_template_format", "plain")
        body_value = template_settings.get("email_template", "")
        if template_format == "html":
            self.template_text.setHtml(self._html_for_editor_preview(body_value))
        else:
            self.template_text.setPlainText(body_value)
        right_layout.addWidget(self.template_text, 1)

        attachments_title = QLabel("Archivos adjuntos de la plantilla")
        attachments_title.setObjectName("sectionTitle")
        right_layout.addWidget(attachments_title)
        attachments_help = QLabel(
            "Registro de archivos cargados: muestra ruta original y copia segura en ~\\.precita\\."
        )
        attachments_help.setWordWrap(True)
        right_layout.addWidget(attachments_help)
        self.attachments_list = QListWidget()
        self.attachments_list.setMinimumHeight(120)
        self.attachments_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        right_layout.addWidget(self.attachments_list)
        self._refresh_attachments_log()
        self.bold_btn = bold_btn
        self.italic_btn = italic_btn
        self.underline_btn = underline_btn
        self.bullet_btn = bullet_btn
        self._format_shortcuts = []
        self._setup_format_shortcuts()
        self.template_text.currentCharFormatChanged.connect(
            self._update_format_buttons_state
        )
        self.template_text.cursorPositionChanged.connect(
            self._update_format_buttons_state
        )
        self._update_format_buttons_state()

        bold_btn.clicked.connect(lambda: self._toggle_bold())
        italic_btn.clicked.connect(lambda: self._toggle_italic())
        underline_btn.clicked.connect(lambda: self._toggle_underline())
        bullet_btn.clicked.connect(self._toggle_bullets)
        clear_btn.clicked.connect(self._clear_formatting)
        attach_btn.clicked.connect(self._pick_attachments)
        
        outer.addWidget(card)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        cancel_btn = QPushButton("Cancelar")
        save_btn = QPushButton("Guardar plantilla")
        save_btn.setObjectName("primaryButton")
        save_btn.clicked.connect(self.save_template)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(save_btn)
        outer.addLayout(button_layout)
    
    def save_template(self):
        """Guardar la plantilla en la BD."""
        subject_template = self.subject_input.text().strip()
        if not subject_template:
            QMessageBox.warning(self, "Error", "El asunto no puede estar vacío.")
            return

        editor_template = self.template_text.toHtml()
        if not self.template_text.toPlainText().strip():
            QMessageBox.warning(self, "Error", "El cuerpo del mensaje no puede estar vacío.")
            return
        template = self._html_for_storage(editor_template)

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?) '
            'ON CONFLICT(key) DO UPDATE SET value = excluded.value',
            ("email_template", template)
        )
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?) '
            'ON CONFLICT(key) DO UPDATE SET value = excluded.value',
            ("email_subject_template", subject_template)
        )
        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?) '
            'ON CONFLICT(key) DO UPDATE SET value = excluded.value',
            ("email_template_format", "html")
        )
        payload_bytes = _template_payload_size_bytes(template, self.attachments)
        if payload_bytes > MAX_TEMPLATE_PAYLOAD_BYTES:
            QMessageBox.warning(
                self,
                "Límite excedido",
                "El tamaño total de adjuntos + cuerpo supera 16 MB.",
            )
            return

        cursor.execute(
            'INSERT INTO settings (key, value) VALUES (?, ?) '
            'ON CONFLICT(key) DO UPDATE SET value = excluded.value',
            ("email_template_attachments", json.dumps(self.attachments, ensure_ascii=False))
        )
        conn.commit()
        conn.close()
        self.accept()

    def _toggle_bold(self):
        self.template_text.setFontWeight(
            QFont.Weight.Normal
            if self.template_text.fontWeight() > QFont.Weight.Normal
            else QFont.Weight.Bold
        )
        self._update_format_buttons_state()

    def _toggle_italic(self):
        self.template_text.setFontItalic(not self.template_text.fontItalic())
        self._update_format_buttons_state()

    def _toggle_underline(self):
        self.template_text.setFontUnderline(not self.template_text.fontUnderline())
        self._update_format_buttons_state()

    def _toggle_bullets(self):
        cursor = self.template_text.textCursor()
        if cursor.currentList():
            block_fmt = cursor.blockFormat()
            block_fmt.setObjectIndex(-1)
            cursor.setBlockFormat(block_fmt)
            self._update_format_buttons_state()
            return
        list_fmt = QTextListFormat()
        list_fmt.setStyle(QTextListFormat.Style.ListDisc)
        cursor.createList(list_fmt)
        self._update_format_buttons_state()

    def _clear_formatting(self):
        cursor = self.template_text.textCursor()
        if not cursor.hasSelection():
            plain = self.template_text.toPlainText()
            self.template_text.setPlainText(plain)
            reset_cursor = self.template_text.textCursor()
            reset_cursor.movePosition(QTextCursor.MoveOperation.End)
            self.template_text.setTextCursor(reset_cursor)
            self._update_format_buttons_state()
            return

        sel_start = min(cursor.selectionStart(), cursor.selectionEnd())
        sel_end = max(cursor.selectionStart(), cursor.selectionEnd())
        if sel_start == sel_end:
            return

        plain = self.template_text.toPlainText()
        selected_plain = plain[sel_start:sel_end]

        replacement_cursor = self.template_text.textCursor()
        replacement_cursor.setPosition(sel_start)
        replacement_cursor.setPosition(sel_end, QTextCursor.MoveMode.KeepAnchor)
        replacement_cursor.insertText(selected_plain, QTextCharFormat())

        first_block = self.template_text.document().findBlock(sel_start)
        last_block_pos = max(sel_start, sel_end - 1)
        last_block = self.template_text.document().findBlock(last_block_pos)
        block = first_block
        while block.isValid():
            block_cursor = QTextCursor(block)
            block_cursor.mergeBlockFormat(QTextBlockFormat())
            clean_block_fmt = block_cursor.blockFormat()
            clean_block_fmt.setObjectIndex(-1)
            block_cursor.setBlockFormat(clean_block_fmt)
            if block == last_block:
                break
            block = block.next()

        selection_cursor = self.template_text.textCursor()
        selection_cursor.setPosition(sel_start)
        selection_cursor.setPosition(sel_start + len(selected_plain), QTextCursor.MoveMode.KeepAnchor)
        self.template_text.setTextCursor(selection_cursor)
        self._update_format_buttons_state()

    def _setup_format_shortcuts(self):
        shortcuts = (
            ("Ctrl+B", self._toggle_bold),
            ("Ctrl+N", self._toggle_bold),
            ("Ctrl+I", self._toggle_italic),
            ("Ctrl+U", self._toggle_underline),
            ("Ctrl+S", self._toggle_underline),
        )
        for sequence, callback in shortcuts:
            shortcut = QShortcut(QKeySequence(sequence), self.template_text)
            shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
            shortcut.activated.connect(callback)
            self._format_shortcuts.append(shortcut)

    def _update_format_buttons_state(self):
        current_format = self.template_text.currentCharFormat()
        cursor = self.template_text.textCursor()
        self.bold_btn.setChecked(current_format.fontWeight() > QFont.Weight.Normal)
        self.italic_btn.setChecked(current_format.fontItalic())
        self.underline_btn.setChecked(current_format.fontUnderline())
        self.bullet_btn.setChecked(cursor.currentList() is not None)

    def _resize_selected_image_by_factor(self, src, scale_factor):
        current = self._extract_image_width_percent(src) or 100
        new_width = max(5, min(100, int(round(current * scale_factor))))
        self._set_image_width_percent(src, new_width)

    def _show_editor_context_menu(self, pos: QPoint):
        menu = self.template_text.createStandardContextMenu()
        cursor = self.template_text.cursorForPosition(pos)
        char_format = cursor.charFormat()
        if char_format.isImageFormat():
            image_format = char_format.toImageFormat()
            src = str(image_format.name() or "").strip()
            if src.startswith("cid:") or src.startswith("data:image/"):
                menu.addSeparator()
                image_menu = menu.addMenu("Imagen integrada")
                size_50_action = image_menu.addAction("Ancho 50%")
                size_75_action = image_menu.addAction("Ancho 75%")
                size_100_action = image_menu.addAction("Ancho 100%")
                size_custom_action = image_menu.addAction("Ancho personalizado…")
                image_menu.addSeparator()
                inline_action = image_menu.addAction("En texto")
                behind_action = image_menu.addAction("Tras texto (fondo)")
                image_menu.addSeparator()
                center_action = image_menu.addAction("Centrar")
                left_action = image_menu.addAction("Alinear izquierda")
                right_action = image_menu.addAction("Alinear derecha")
                image_menu.addSeparator()
                move_action = image_menu.addAction("Mover (X/Y)…")

                chosen = menu.exec(self.template_text.mapToGlobal(pos))
                if chosen == size_50_action:
                    self._set_image_width_percent(src, 50)
                elif chosen == size_75_action:
                    self._set_image_width_percent(src, 75)
                elif chosen == size_100_action:
                    self._set_image_width_percent(src, 100)
                elif chosen == size_custom_action:
                    self._set_image_width_percent_custom(src)
                elif chosen == inline_action:
                    self._set_image_layout_mode(src, mode="inline")
                elif chosen == behind_action:
                    self._set_image_layout_mode(src, mode="behind")
                elif chosen == center_action:
                    self._set_image_alignment(src, align="center")
                elif chosen == left_action:
                    self._set_image_alignment(src, align="left")
                elif chosen == right_action:
                    self._set_image_alignment(src, align="right")
                elif chosen == move_action:
                    self._move_background_image(src)
                return
        menu.exec(self.template_text.mapToGlobal(pos))

    def _html_for_editor_preview(self, html):
        if not html or "<img" not in html:
            return html
        cid_to_path = {}
        for item in self.attachments:
            if item.get("inline") and item.get("cid"):
                cid_to_path[str(item.get("cid"))] = str(item.get("path", ""))
        if not cid_to_path:
            return html
        pattern = re.compile(r"<img\b[^>]*>", re.IGNORECASE | re.DOTALL)

        def _replace_tag(match):
            tag = match.group(0)
            src_match = re.search(r"\bsrc=(['\"])(.*?)\1", tag, re.IGNORECASE | re.DOTALL)
            if not src_match:
                return tag
            src_value = src_match.group(2).strip()
            if not src_value.startswith("cid:"):
                return tag
            cid_value = src_value[4:]
            stored_path = Path(cid_to_path.get(cid_value, "")).expanduser()
            if not stored_path.is_file():
                return tag
            data_uri = self._build_data_uri(stored_path)
            if not data_uri:
                return tag
            if "data-precita-cid=" in tag:
                return tag[:src_match.start(2)] + data_uri + tag[src_match.end(2):]
            updated = tag[:src_match.start(2)] + data_uri + tag[src_match.end(2):]
            return updated[:-1] + f' data-precita-cid="{cid_value}">'

        return pattern.sub(_replace_tag, html)

    def _html_for_storage(self, html):
        if not html or "<img" not in html:
            return html
        pattern = re.compile(r"<img\b[^>]*>", re.IGNORECASE | re.DOTALL)

        def _replace_tag(match):
            tag = match.group(0)
            cid_match = re.search(r"\bdata-precita-cid=(['\"])(.*?)\1", tag, re.IGNORECASE | re.DOTALL)
            if not cid_match:
                return tag
            cid_value = cid_match.group(2).strip()
            if not cid_value:
                return tag
            src_match = re.search(r"\bsrc=(['\"])(.*?)\1", tag, re.IGNORECASE | re.DOTALL)
            if src_match:
                return tag[:src_match.start(2)] + f"cid:{cid_value}" + tag[src_match.end(2):]
            return tag[:-1] + f' src="cid:{cid_value}">'

        return pattern.sub(_replace_tag, html)

    def _build_data_uri(self, image_path):
        try:
            mime_type, _ = mimetypes.guess_type(str(image_path))
            if not (mime_type and mime_type.startswith("image/")):
                return ""
            with open(image_path, "rb") as file_obj:
                encoded = base64.b64encode(file_obj.read()).decode("ascii")
            return f"data:{mime_type};base64,{encoded}"
        except Exception:
            return ""

    def _set_image_width_percent_custom(self, src):
        current = self._extract_image_width_percent(src) or 100
        value, ok = QInputDialog.getInt(
            self,
            "Ancho de imagen",
            "Porcentaje de ancho (10-100):",
            current,
            10,
            100,
            1,
        )
        if ok:
            self._set_image_width_percent(src, value)

    def _set_image_width_percent(self, src, width_percent):
        width = max(10, min(100, int(width_percent)))
        def _updater(style, attrs):
            mode = attrs.get("data-precita-mode", "inline")
            align = attrs.get("data-precita-align", "left")
            left = attrs.get("data-precita-left", "0")
            top = attrs.get("data-precita-top", "0")
            if mode == "behind":
                new_style = (
                    f"position:absolute; left:{left}px; top:{top}px; "
                    f"width:{width}%; opacity:0.28; z-index:0;"
                )
            else:
                margin_map = {
                    "left": "margin: 8px 12px 8px 0;",
                    "center": "margin: 8px auto;",
                    "right": "margin: 8px 0 8px auto;",
                }
                margin_css = margin_map.get(align, margin_map["left"])
                new_style = (
                    f"display:block; max-width: {width}%; height:auto; {margin_css}"
                )
            return new_style
        self._update_image_by_src(src, _updater)

    def _set_image_layout_mode(self, src, mode):
        mode = "behind" if mode == "behind" else "inline"
        def _updater(style, attrs):
            width = self._extract_image_width_percent(src) or 100
            align = attrs.get("data-precita-align", "left")
            left = attrs.get("data-precita-left", "0")
            top = attrs.get("data-precita-top", "0")
            if mode == "behind":
                return (
                    f"position:absolute; left:{left}px; top:{top}px; "
                    f"width:{width}%; opacity:0.28; z-index:0;"
                )
            margin_map = {
                "left": "margin: 8px 12px 8px 0;",
                "center": "margin: 8px auto;",
                "right": "margin: 8px 0 8px auto;",
            }
            margin_css = margin_map.get(align, margin_map["left"])
            return f"display:block; max-width: {width}%; height:auto; {margin_css}"
        self._update_image_by_src(src, _updater, extra_attrs={"data-precita-mode": mode})

    def _set_image_alignment(self, src, align):
        if align not in {"left", "center", "right"}:
            return
        def _updater(style, attrs):
            mode = attrs.get("data-precita-mode", "inline")
            width = self._extract_image_width_percent(src) or 100
            left = attrs.get("data-precita-left", "0")
            top = attrs.get("data-precita-top", "0")
            if mode == "behind":
                if align == "left":
                    left_px = "0"
                elif align == "center":
                    left_px = "30"
                else:
                    left_px = "60"
                return (
                    f"position:absolute; left:{left_px}px; top:{top}px; "
                    f"width:{width}%; opacity:0.28; z-index:0;"
                )
            margin_map = {
                "left": "margin: 8px 12px 8px 0;",
                "center": "margin: 8px auto;",
                "right": "margin: 8px 0 8px auto;",
            }
            margin_css = margin_map.get(align, margin_map["left"])
            return f"display:block; max-width: {width}%; height:auto; {margin_css}"
        self._update_image_by_src(src, _updater, extra_attrs={"data-precita-align": align})

    def _move_background_image(self, src):
        x, ok_x = QInputDialog.getInt(self, "Mover imagen (X)", "Desplazamiento X (px):", 0, -2000, 2000, 1)
        if not ok_x:
            return
        y, ok_y = QInputDialog.getInt(self, "Mover imagen (Y)", "Desplazamiento Y (px):", 0, -2000, 2000, 1)
        if not ok_y:
            return
        def _updater(style, attrs):
            width = self._extract_image_width_percent(src) or 100
            return (
                f"position:absolute; left:{x}px; top:{y}px; "
                f"width:{width}%; opacity:0.28; z-index:0;"
            )
        self._update_image_by_src(
            src,
            _updater,
            extra_attrs={
                "data-precita-mode": "behind",
                "data-precita-left": str(x),
                "data-precita-top": str(y),
            },
        )

    def _extract_image_width_percent(self, src):
        html = self.template_text.toHtml()
        pattern = re.compile(
            r"<img\b(?=[^>]*\bsrc=(['\"])%s\1)[^>]*>"
            % re.escape(src),
            re.IGNORECASE | re.DOTALL,
        )
        match = pattern.search(html)
        if not match:
            return None
        tag = match.group(0)
        style_match = re.search(r"\bstyle=(['\"])(.*?)\1", tag, re.IGNORECASE | re.DOTALL)
        if not style_match:
            return None
        style = style_match.group(2)
        width_match = re.search(r"(?:max-width|width)\s*:\s*(\d+)\s*%", style, re.IGNORECASE)
        if not width_match:
            return None
        try:
            return int(width_match.group(1))
        except Exception:
            return None

    def _update_image_by_src(self, src, style_updater, extra_attrs=None):
        html = self.template_text.toHtml()
        pattern = re.compile(
            r"<img\b(?=[^>]*\bsrc=(['\"])%s\1)[^>]*>"
            % re.escape(src),
            re.IGNORECASE | re.DOTALL,
        )
        match = pattern.search(html)
        if not match:
            return
        tag = match.group(0)

        attrs = {}
        for key in ("data-precita-mode", "data-precita-align", "data-precita-left", "data-precita-top"):
            attr_match = re.search(rf"\b{key}=(['\"])(.*?)\1", tag, re.IGNORECASE | re.DOTALL)
            if attr_match:
                attrs[key] = attr_match.group(2)
        if extra_attrs:
            attrs.update(extra_attrs)

        style_match = re.search(r"\bstyle=(['\"])(.*?)\1", tag, re.IGNORECASE | re.DOTALL)
        old_style = style_match.group(2) if style_match else ""
        new_style = style_updater(old_style, attrs).strip()

        if style_match:
            updated_tag = tag[:style_match.start(2)] + new_style + tag[style_match.end(2):]
        else:
            updated_tag = tag[:-1] + f' style="{new_style}">'

        for key, value in attrs.items():
            attr_pattern = re.compile(rf"\b{key}=(['\"])(.*?)\1", re.IGNORECASE | re.DOTALL)
            if attr_pattern.search(updated_tag):
                updated_tag = attr_pattern.sub(f'{key}="{value}"', updated_tag, count=1)
            else:
                updated_tag = updated_tag[:-1] + f' {key}="{value}">'

        updated_html = html[:match.start()] + updated_tag + html[match.end():]
        self.template_text.setHtml(updated_html)
        self.template_text.select_image_by_src(src)

    def _store_attachment_in_precita(self, source_path):
        source = Path(source_path).expanduser().resolve()
        digest = hashlib.sha256()
        with open(source, "rb") as file_obj:
            while True:
                chunk = file_obj.read(1024 * 1024)
                if not chunk:
                    break
                digest.update(chunk)
        file_hash = digest.hexdigest()[:16]
        target_name = f"{file_hash}_{source.name}"
        target_path = ATTACHMENTS_DIR / target_name
        if not target_path.exists():
            shutil.copy2(source, target_path)
        return {
            "path": str(target_path),
            "name": source.name,
            "original_path": str(source),
        }

    def _add_attachment_item(self, attachment_item):
        stored_path = str(attachment_item.get("path", "")).strip()
        display_name = str(attachment_item.get("name", "")).strip() or Path(stored_path).name
        original_path = str(attachment_item.get("original_path", "")).strip()
        is_inline = bool(attachment_item.get("inline", False))
        cid_value = str(attachment_item.get("cid", "")).strip() if is_inline else ""
        if not stored_path:
            return False
        normalized_path = str(Path(stored_path).expanduser().resolve())
        normalized_original_path = ""
        if original_path:
            try:
                normalized_original_path = str(Path(original_path).expanduser().resolve())
            except Exception:
                normalized_original_path = original_path
        if is_inline:
            if not cid_value:
                return False
            already_present = any(
                item.get("path") == normalized_path and item.get("cid") == cid_value
                for item in self.attachments
            )
        else:
            already_present = any(
                item.get("path") == normalized_path and not item.get("inline", False)
                for item in self.attachments
            )
        if already_present:
            return False
        item = {
            "path": normalized_path,
            "name": display_name,
            "original_path": normalized_original_path,
            "inline": is_inline,
            "cid": cid_value,
        }
        self.attachments.append(item)
        self._refresh_attachments_log()
        return True

    def _format_attachment_log_line(self, item):
        item_kind = "Imagen integrada" if item.get("inline", False) else "Adjunto"
        name = item.get("name", Path(item.get("path", "")).name)
        original_path = str(item.get("original_path", "")).strip() or "(no disponible)"
        secure_path = str(item.get("path", "")).strip() or "(sin copia segura)"
        cid_hint = f" — cid:{item.get('cid', '')}" if item.get("inline", False) else ""
        return (
            f"[{item_kind}] {name}{cid_hint}\n"
            f"Origen: {original_path}\n"
            f"Seguro: {secure_path}"
        )

    def _refresh_attachments_log(self):
        if not hasattr(self, "attachments_list"):
            return
        self.attachments_list.clear()
        for item in self.attachments:
            self.attachments_list.addItem(self._format_attachment_log_line(item))

    def _pick_attachments(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Adjuntar archivos",
            str(Path.home()),
            "Todos los archivos (*.*)"
        )
        rejected = []
        added_items = []
        for file_path in files:
            is_valid, reason = _validate_attachment_security(file_path)
            if not is_valid:
                rejected.append(f"- {Path(file_path).name}: {reason}")
                continue
            stored_item = self._store_attachment_in_precita(file_path)
            if self._add_attachment_item(stored_item):
                added_items.append(stored_item)

        payload_bytes = _template_payload_size_bytes(self.template_text.toHtml(), self.attachments)
        if payload_bytes > MAX_TEMPLATE_PAYLOAD_BYTES:
            added_paths = {str(Path(item.get("path", "")).expanduser().resolve()) for item in added_items}
            self.attachments = [
                item for item in self.attachments
                if str(Path(item.get("path", "")).expanduser().resolve()) not in added_paths
            ]
            self._refresh_attachments_log()
            for item in added_items:
                stored_path = Path(item.get("path", "")).expanduser()
                if stored_path.exists():
                    try:
                        stored_path.unlink()
                    except Exception:
                        pass
            rejected.append("- Se excede el límite de 16 MB al añadir estos adjuntos.")

        if rejected:
            QMessageBox.warning(
                self,
                "Adjuntos bloqueados",
                "Algunos archivos no se pudieron adjuntar:\n\n" + "\n".join(rejected),
            )

    def _pick_inline_images(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Insertar imágenes integradas",
            str(Path.home()),
            "Imágenes (*.png *.jpg *.jpeg *.gif *.bmp *.webp)"
        )
        rejected = []
        added_items = []
        for file_path in files:
            is_valid, reason = _validate_attachment_security(file_path)
            if not is_valid:
                rejected.append(f"- {Path(file_path).name}: {reason}")
                continue
            mime_type, _ = mimetypes.guess_type(str(file_path))
            if not (mime_type and mime_type.startswith("image/")):
                rejected.append(f"- {Path(file_path).name}: no es un formato de imagen válido.")
                continue
            stored_item = self._store_attachment_in_precita(file_path)
            cid_value = f"precita-inline-{uuid4().hex[:20]}"
            stored_item["inline"] = True
            stored_item["cid"] = cid_value
            if self._add_attachment_item(stored_item):
                added_items.append(stored_item)
                data_uri = self._build_data_uri(Path(stored_item.get("path", "")).expanduser())
                if not data_uri:
                    rejected.append(f"- {Path(file_path).name}: no se pudo renderizar en el editor.")
                    continue
                self.template_text.insertHtml(
                    f'<img src="{data_uri}" data-precita-cid="{cid_value}" alt="{stored_item.get("name", "imagen")}" style="display:block; max-width: 45%; height:auto; margin: 8px 12px 8px 0;"><br>'
                )

        payload_bytes = _template_payload_size_bytes(
            self._html_for_storage(self.template_text.toHtml()),
            self.attachments
        )
        if payload_bytes > MAX_TEMPLATE_PAYLOAD_BYTES:
            added_signatures = {
                (str(Path(item.get("path", "")).expanduser().resolve()), item.get("cid", ""))
                for item in added_items
            }
            self.attachments = [
                item for item in self.attachments
                if (
                    str(Path(item.get("path", "")).expanduser().resolve()),
                    item.get("cid", "")
                ) not in added_signatures
            ]
            self._refresh_attachments_log()
            for item in added_items:
                stored_path = Path(item.get("path", "")).expanduser()
                if stored_path.exists():
                    try:
                        stored_path.unlink()
                    except Exception:
                        pass
            rejected.append("- Se excede el límite de 16 MB al añadir estas imágenes integradas.")

        if rejected:
            QMessageBox.warning(
                self,
                "Imágenes integradas bloqueadas",
                "Algunas imágenes no se pudieron insertar:\n\n" + "\n".join(rejected),
            )

    def _remove_selected_attachment(self):
        row = self.attachments_list.currentRow()
        if row < 0:
            return
        self.attachments_list.takeItem(row)
        removed = self.attachments.pop(row)
        stored_path = Path(removed.get("path", "")).expanduser()
        if stored_path.exists():
            try:
                stored_path.unlink()
            except Exception:
                pass

    def _clear_attachments(self):
        for item in self.attachments:
            stored_path = Path(item.get("path", "")).expanduser()
            if stored_path.exists():
                try:
                    stored_path.unlink()
                except Exception:
                    pass
        self.attachments_list.clear()
        self.attachments = []

# ============================================================================
# TABLA DE CONTACTOS (ordenación coherente con COLLATE NOCASE en SQLite)
# ============================================================================


class _CaseInsensitiveTableWidgetItem(QTableWidgetItem):
    """Compara textos sin distinguir mayúsculas al ordenar columnas."""

    def __lt__(self, other):
        if other is None:
            return False
        if not isinstance(other, QTableWidgetItem):
            return NotImplemented
        return (self.text() or "").casefold() < (other.text() or "").casefold()


# ============================================================================
# DIÁLOGO DE NUEVO CONTACTO
# ============================================================================

class NewContactDialog(QDialog):
    """Diálogo para agregar un nuevo contacto."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Añadir contacto — PreCita — {get_version()}")
        self.setMinimumWidth(400)
        self.resize(440, 360)
        
        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(12)
        
        card = QFrame()
        card.setObjectName("panelCard")
        form = QFormLayout(card)
        form.setContentsMargins(16, 16, 16, 16)
        form.setSpacing(12)
        form.setHorizontalSpacing(14)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        
        self.first_name_input = QLineEdit()
        self.last_name_input = QLineEdit()
        self.email_input = QLineEdit()
        self.phone_input = QLineEdit()
        self.phone_input.setValidator(_PHONE_DIGITS_VALIDATOR)
        self.first_name_input.setPlaceholderText("*")
        self.last_name_input.setPlaceholderText("*")
        self.email_input.setPlaceholderText("*")
        self.phone_input.setPlaceholderText("Opcional, solo números")
        
        form.addRow("Nombre", self.first_name_input)
        form.addRow("Apellido(s)", self.last_name_input)
        form.addRow("Correo", self.email_input)
        form.addRow("Teléfono", self.phone_input)
        outer.addWidget(card)

        info_label = QLabel(
            "Puede ver los contactos guardados en "
            "<a href='open_contacts'>Gestión de contactos</a>. "
            "Todos los contactos se almacenan localmente. "
            "Entre en <a href='open_help'>Ayuda</a> para más información."
        )
        info_label.setWordWrap(True)
        info_label.setTextFormat(Qt.TextFormat.RichText)
        info_label.setOpenExternalLinks(False)
        info_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        info_label.linkActivated.connect(self._open_related_section)
        outer.addWidget(info_label)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        cancel_btn = QPushButton("Cancelar")
        save_btn = QPushButton("Guardar contacto")
        save_btn.setObjectName("primaryButton")
        save_btn.clicked.connect(self.save_contact)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(save_btn)
        outer.addLayout(button_layout)

    def _open_related_section(self, destination):
        """Abrir secciones relacionadas desde enlaces del diálogo."""
        parent_window = self.parent()
        if not parent_window:
            return
        self.reject()
        if destination == "open_contacts" and hasattr(parent_window, "open_contacts_dialog"):
            parent_window.open_contacts_dialog()
        elif destination == "open_help" and hasattr(parent_window, "open_help_dialog"):
            parent_window.open_help_dialog()
    
    def save_contact(self):
        """Guardar contacto en la BD."""
        first_name = self.first_name_input.text().strip()
        last_name = self.last_name_input.text().strip()
        email = self.email_input.text().strip()
        phone = self.phone_input.text().strip()
        if phone and not phone.isdigit():
            QMessageBox.warning(
                self, "Error", "El teléfono sólo puede contener números."
            )
            return

        if not first_name or not last_name or not email:
            QMessageBox.warning(
                self, "Error", "Debe rellenar los campos con asteriscos."
            )
            return
        if not is_plausible_email(email):
            QMessageBox.warning(
                self,
                "Error",
                "El correo no tiene un formato válido. "
                "Ejemplo: nombre@dominio.com",
            )
            return
        if not is_known_mail(email):
            confirm = QMessageBox(self)
            confirm.setIcon(QMessageBox.Icon.Warning)
            confirm.setWindowTitle("Dominio de correo no reconocido")
            confirm.setText(
                f"El dominio del correo ({email.split('@')[-1].lower()}) no coincide con los dominios más comunes.\n"
                "Podría haber un error de escritura.\n\n"
                "¿Desea modificar el correo o guardar el contacto de todos modos?"
            )
            modificar_btn = confirm.addButton("Modificar", QMessageBox.ButtonRole.RejectRole)
            continuar_btn = confirm.addButton("Continuar", QMessageBox.ButtonRole.AcceptRole)
            modificar_btn.setObjectName("primaryButton")
            modificar_btn.style().unpolish(modificar_btn)
            modificar_btn.style().polish(modificar_btn)
            confirm.setDefaultButton(modificar_btn)
            confirm.exec()
            if confirm.clickedButton() == modificar_btn:
                self.email_input.setFocus()
                self.email_input.selectAll()
                return
        
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute(
                'INSERT INTO contacts (first_name, last_name, email, phone) VALUES (?, ?, ?, ?)',
                (first_name, last_name, email, phone)
            )
            conn.commit()
            conn.close()
            self.accept()
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Error", "Este email ya existe.")


class ContactsDialog(QDialog):
    """Ventana emergente de contactos con estilo de sección antigua."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Gestión de contactos — PreCita — {get_version()}")
        self.setMinimumSize(860, 520)
        self.resize(960, 620)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(12)

        contacts_card = QFrame()
        contacts_card.setObjectName("panelCard")
        contacts_inner = QVBoxLayout(contacts_card)
        contacts_inner.setContentsMargins(14, 12, 14, 12)
        contacts_inner.setSpacing(8)
        contacts_label = QLabel("Contactos")
        contacts_label.setObjectName("sectionTitle")
        contacts_inner.addWidget(contacts_label)

        self.contacts_table = QTableWidget()
        self.contacts_table.setColumnCount(4)
        self.contacts_table.setHorizontalHeaderLabels(
            ["Nombre", "Apellido(s)", "Correo", "Teléfono"]
        )
        self.contacts_table.setAlternatingRowColors(True)
        self.contacts_table.setShowGrid(False)
        self.contacts_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.contacts_table.verticalHeader().setVisible(False)
        self.contacts_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.contacts_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self.contacts_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.Stretch
        )
        self.contacts_table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.ResizeToContents
        )
        self.contacts_table.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self.contacts_table.verticalHeader().setDefaultSectionSize(40)
        self.contacts_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.contacts_table.customContextMenuRequested.connect(self._contacts_context_menu)
        self.contacts_table.setSortingEnabled(True)
        self.contacts_table.sortByColumn(1, Qt.SortOrder.AscendingOrder)
        self.contacts_table.horizontalHeader().setSortIndicatorShown(False)
        contacts_inner.addWidget(self.contacts_table)
        outer.addWidget(contacts_card, 1)

        button_layout = QHBoxLayout()
        new_btn = QPushButton("Añadir contacto")
        new_btn.clicked.connect(self.open_new_contact_dialog)
        close_btn = QPushButton("Cerrar")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(new_btn)
        button_layout.addStretch()
        button_layout.addWidget(close_btn)
        outer.addLayout(button_layout)

        self.load_contacts()

    def load_contacts(self):
        hdr = self.contacts_table.horizontalHeader()
        sort_col = hdr.sortIndicatorSection()
        sort_order = hdr.sortIndicatorOrder()
        if sort_col < 0:
            sort_col = 1
            sort_order = Qt.SortOrder.AscendingOrder

        self.contacts_table.setSortingEnabled(False)
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT id, first_name, last_name, email, phone FROM contacts
            ORDER BY last_name COLLATE NOCASE ASC, first_name COLLATE NOCASE ASC
            """
        )
        contacts = cursor.fetchall()
        conn.close()
        self.contacts_table.setRowCount(len(contacts))

        for row, (cid, first_name, last_name, email, phone) in enumerate(contacts):
            name_item = _CaseInsensitiveTableWidgetItem(first_name)
            name_item.setData(Qt.ItemDataRole.UserRole, cid)
            self.contacts_table.setItem(row, 0, name_item)
            self.contacts_table.setItem(
                row, 1, _CaseInsensitiveTableWidgetItem(last_name or "")
            )
            self.contacts_table.setItem(row, 2, _CaseInsensitiveTableWidgetItem(email))
            self.contacts_table.setItem(
                row, 3, _CaseInsensitiveTableWidgetItem(phone or "-")
            )

        self.contacts_table.setSortingEnabled(True)
        self.contacts_table.sortByColumn(sort_col, sort_order)

    def _contacts_context_menu(self, pos):
        row = self.contacts_table.rowAt(pos.y())
        if row < 0:
            return
        name_item = self.contacts_table.item(row, 0)
        if not name_item:
            return
        cid = name_item.data(Qt.ItemDataRole.UserRole)
        if cid is None:
            return
        self.contacts_table.selectRow(row)
        cid = int(cid)
        menu = QMenu(self)
        menu.addAction("Editar nombre del contacto", lambda: self._edit_contact_first_name(cid))
        menu.addAction(
            "Editar apellido(s) del contacto", lambda: self._edit_contact_last_name(cid)
        )
        menu.addAction("Editar correo del contacto", lambda: self._edit_contact_email(cid))
        menu.addAction("Editar teléfono del contacto", lambda: self._edit_contact_phone(cid))
        menu.addSeparator()
        delete_action = QAction("Eliminar contacto", menu)
        delete_action.setObjectName("precitaMenuDelete")
        delete_action.setProperty("precitaDestructive", True)
        delete_action.triggered.connect(lambda: self._delete_contact(cid))
        menu.addAction(delete_action)
        menu.exec(self.contacts_table.viewport().mapToGlobal(pos))

    def _edit_contact_first_name(self, contact_id):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT first_name FROM contacts WHERE id = ?', (contact_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return
        current = row[0]
        text, ok = QInputDialog.getText(
            self, "Editar nombre", "Nombre del contacto:", text=current
        )
        if not ok:
            return
        first_name = text.strip()
        if not first_name:
            QMessageBox.warning(self, "Error", "El nombre no puede estar vacío.")
            return
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            'UPDATE contacts SET first_name = ? WHERE id = ?', (first_name, contact_id)
        )
        conn.commit()
        conn.close()
        self._refresh_after_contact_update("✓ Nombre de contacto actualizado")

    def _edit_contact_last_name(self, contact_id):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT last_name FROM contacts WHERE id = ?', (contact_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return
        current = row[0] or ""
        text, ok = QInputDialog.getText(
            self, "Editar apellido(s)", "Apellido(s) del contacto:", text=current
        )
        if not ok:
            return
        last_name = text.strip()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            'UPDATE contacts SET last_name = ? WHERE id = ?', (last_name, contact_id)
        )
        conn.commit()
        conn.close()
        self._refresh_after_contact_update("✓ Apellido(s) de contacto actualizado(s)")

    def _edit_contact_email(self, contact_id):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT email FROM contacts WHERE id = ?', (contact_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return
        current = row[0]
        text, ok = QInputDialog.getText(
            self, "Editar correo", "Correo electrónico:", text=current
        )
        if not ok:
            return
        email = text.strip()
        if not email:
            QMessageBox.warning(self, "Error", "El correo no puede estar vacío.")
            return
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('UPDATE contacts SET email = ? WHERE id = ?', (email, contact_id))
            conn.commit()
            conn.close()
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Error", "Ese correo ya está registrado en otro contacto.")
            return
        self._refresh_after_contact_update("✓ Correo de contacto actualizado")

    def _edit_contact_phone(self, contact_id):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT phone FROM contacts WHERE id = ?', (contact_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return
        current = row[0] or ""
        dlg = QInputDialog(self)
        dlg.setWindowTitle("Editar teléfono")
        dlg.setLabelText("Teléfono (opcional, solo números):")
        dlg.setTextValue(current)
        le = dlg.findChild(QLineEdit)
        if le is not None:
            le.setValidator(_PHONE_DIGITS_VALIDATOR)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        phone = dlg.textValue().strip()
        if phone and not phone.isdigit():
            QMessageBox.warning(self, "Error", "El teléfono sólo puede contener números.")
            return
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('UPDATE contacts SET phone = ? WHERE id = ?', (phone, contact_id))
        conn.commit()
        conn.close()
        self._refresh_after_contact_update("✓ Teléfono de contacto actualizado")

    def _delete_contact(self, contact_id):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            'SELECT first_name, last_name FROM contacts WHERE id = ?', (contact_id,)
        )
        row = cursor.fetchone()
        if not row:
            conn.close()
            return
        display = contact_full_name(row[0], row[1]) or "contacto"
        conn.close()
        reply = QMessageBox.question(
            self,
            "Eliminar contacto",
            f"¿Eliminar el contacto «{display}»?\n"
            "Las citas vinculadas quedarán sin contacto asignado.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            'UPDATE appointments SET contact_id = NULL WHERE contact_id = ?', (contact_id,)
        )
        cursor.execute('DELETE FROM contacts WHERE id = ?', (contact_id,))
        conn.commit()
        conn.close()
        self._refresh_after_contact_update(f"✓ Contacto eliminado: {display}")

    def open_new_contact_dialog(self):
        dialog = NewContactDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._refresh_after_contact_update("✓ Nuevo contacto agregado")

    def _refresh_after_contact_update(self, message):
        self.load_contacts()
        mw = self.parent()
        if mw and hasattr(mw, "log_message"):
            mw.log_message(message)
        if mw and hasattr(mw, "refresh_tables"):
            mw.refresh_tables()


REMINDER_INTERVAL_CHOICES_SEC = (300, 600, 900, 1800, 3600, 7200, 14400)
REMINDER_INTERVAL_LABELS = (
    "5 minutos",
    "10 minutos",
    "15 minutos",
    "20 minutos",
    "30 minutos",
    "1 hora",
    "2 horas",
)

DEFAULT_THEME = "dark"
DEFAULT_DISPLAY_SCALE_PERCENT = 100
DISPLAY_SCALE_CHOICES_PERCENT = (80, 90, 100, 110, 125, 150)
DEFAULT_WINDOWS_STARTUP = "1"
DEFAULT_WINDOWS_NOTIFICATIONS = "1"
DEFAULT_REMINDER_INTERVAL_SEC = REMINDER_INTERVAL_CHOICES_SEC[0]
MIN_CALENDAR_SYNC_DAYS = 7
MAX_CALENDAR_SYNC_DAYS = 49
DEFAULT_CALENDAR_SYNC_DAYS = 15


class SettingsDialog(QDialog):
    """Preferencias: tema, inicio con Windows, intervalo de revisión, cuenta Google."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Configuración — PreCita — {get_version()}")
        self.setMinimumWidth(680)
        self.resize(760, 620)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(10)

        scroll = QScrollArea()
        scroll.setObjectName("settingsScrollArea")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        outer.addWidget(scroll, 1)

        content = QWidget()
        content.setObjectName("settingsContent")
        scroll.setWidget(content)
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(20, 20, 20, 0)
        content_layout.setSpacing(14)

        appearance = QGroupBox("Apariencia")
        app_form = QFormLayout(appearance)
        app_form.setSpacing(10)
        self.theme_combo = QComboBox()
        self.theme_combo.addItem("Claro", "light")
        self.theme_combo.addItem("Oscuro", "dark")
        current_theme = get_setting("theme", DEFAULT_THEME) or DEFAULT_THEME
        idx = self.theme_combo.findData(current_theme)
        if idx >= 0:
            self.theme_combo.setCurrentIndex(idx)
        self.display_scale_combo = QComboBox()
        for percent in DISPLAY_SCALE_CHOICES_PERCENT:
            self.display_scale_combo.addItem(f"{percent} %", percent)
        current_scale = get_display_scale_percent()
        idx_scale = self.display_scale_combo.findData(current_scale)
        self.display_scale_combo.setCurrentIndex(idx_scale if idx_scale >= 0 else 0)
        appearance_row = QWidget()
        appearance_row_layout = QHBoxLayout(appearance_row)
        appearance_row_layout.setContentsMargins(0, 0, 0, 0)
        appearance_row_layout.setSpacing(8)
        appearance_row_layout.addWidget(QLabel("Tema"))
        appearance_row_layout.addWidget(self.theme_combo)
        appearance_row_layout.addSpacing(12)
        appearance_row_layout.addWidget(QLabel("Tamaño de visualización"))
        appearance_row_layout.addWidget(self.display_scale_combo)
        appearance_row_layout.addStretch(1)
        app_form.addRow(appearance_row)
        content_layout.addWidget(appearance)

        system = QGroupBox("Sistema")
        sys_layout = QVBoxLayout(system)
        self.startup_check = QCheckBox("Iniciar PreCita al arrancar Windows")
        if sys.platform == "win32":
            st = get_setting("windows_startup", DEFAULT_WINDOWS_STARTUP)
            self.startup_check.setChecked(st == "1" or windows_startup_is_enabled())
        else:
            self.startup_check.setChecked(False)
            self.startup_check.setEnabled(False)
            self.startup_check.setToolTip("Solo disponible en Windows.")
        self.notifications_check = QCheckBox("Habilitar notificaciones de PreCita")
        notif_enabled = get_setting(
            "windows_notifications_enabled", DEFAULT_WINDOWS_NOTIFICATIONS
        )
        self.notifications_check.setChecked(str(notif_enabled) == "1")
        if sys.platform != "win32":
            self.notifications_check.setChecked(False)
            self.notifications_check.setEnabled(False)
            self.notifications_check.setToolTip("Solo disponible en Windows.")
        system_row = QWidget()
        system_row_layout = QHBoxLayout(system_row)
        system_row_layout.setContentsMargins(0, 0, 0, 0)
        system_row_layout.setSpacing(16)
        system_row_layout.addWidget(self.startup_check)
        system_row_layout.addWidget(self.notifications_check)
        system_row_layout.addStretch(1)
        sys_layout.addWidget(system_row)
        notif_hint = QLabel(
            "'Iniciar PreCita al arrancar Windows': Recomendado para mejor automatización de la aplicación. "
            "'Habilitar notificaciones de PreCita': PreCita usa notificaciones de Windows tras el envío "
            "automático y exitoso de un correo, mostrando contacto destinatario y hora prevista de la "
            "cita programada."
        )
        notif_hint.setObjectName("dialogHint")
        notif_hint.setWordWrap(True)
        sys_layout.addWidget(notif_hint)
        content_layout.addWidget(system)

        cal = QGroupBox("Calendario y recordatorios")
        cal_form = QFormLayout(cal)
        cal_form.setSpacing(10)
        self.interval_combo = QComboBox()
        for sec, label in zip(REMINDER_INTERVAL_CHOICES_SEC, REMINDER_INTERVAL_LABELS):
            self.interval_combo.addItem(label, sec)
        raw_iv = get_setting(
            "reminder_interval_sec", str(DEFAULT_REMINDER_INTERVAL_SEC)
        )
        try:
            iv = int(raw_iv)
        except (TypeError, ValueError):
            iv = DEFAULT_REMINDER_INTERVAL_SEC
        ix = self.interval_combo.findData(iv)
        self.interval_combo.setCurrentIndex(ix if ix >= 0 else 0)
        hint = QLabel(
            "'Revisar cada': Intervalo entre comprobaciones automáticas de recordatorios por correo "
            "(recomendado el rango de 5-15 minutos). "
            "'Sincronizar próximos días': Cantidad de días a sincronizar desde hoy para Google Calendar "
            f"(mínimo {MIN_CALENDAR_SYNC_DAYS}, máximo {MAX_CALENDAR_SYNC_DAYS}). "
            "Use 'Cambiar calendario de Google' para elegir de forma explícita qué calendario "
            "puede leer PreCita."
        )
        hint.setObjectName("dialogHint")
        hint.setWordWrap(True)
        self.interval_combo.setMinimumWidth(90)
        self.interval_combo.setMaximumWidth(100)
        self.sync_days_input = QLineEdit()
        self.sync_days_input.setPlaceholderText(
            f"{MIN_CALENDAR_SYNC_DAYS}-{MAX_CALENDAR_SYNC_DAYS}"
        )
        self.sync_days_input.setValidator(
            QRegularExpressionValidator(QRegularExpression(r"^\d{1,2}$"))
        )
        self.sync_days_input.setMinimumWidth(48)
        self.sync_days_input.setMaximumWidth(56)
        self.sync_days_input.setText(str(get_calendar_sync_days()))
        schedule_row = QWidget()
        schedule_layout = QHBoxLayout(schedule_row)
        schedule_layout.setContentsMargins(0, 0, 0, 0)
        schedule_layout.setSpacing(8)
        schedule_layout.addWidget(QLabel("Revisar cada"))
        schedule_layout.addWidget(self.interval_combo)
        schedule_layout.addSpacing(12)
        schedule_layout.addWidget(QLabel("Sincronizar próximos días"))
        schedule_layout.addWidget(self.sync_days_input)
        schedule_layout.addStretch(1)
        cal_form.addRow(schedule_row)
        self.calendar_selection_label = QLabel()
        self.calendar_selection_label.setObjectName("dialogHint")
        self._refresh_calendar_selection_label()
        self.calendar_selection_btn = QPushButton("Cambiar calendario de Google…")
        self.calendar_selection_btn.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )
        self.calendar_selection_btn.setMinimumHeight(26)
        self.calendar_selection_btn.setStyleSheet("padding: 4px 10px;")
        self.calendar_selection_btn.clicked.connect(self._change_google_calendar)
        calendar_row = QWidget()
        calendar_row_layout = QHBoxLayout(calendar_row)
        calendar_row_layout.setContentsMargins(0, 0, 0, 0)
        calendar_row_layout.setSpacing(8)
        calendar_sync_title = QLabel("Calendario sincronizado")
        calendar_row_layout.addWidget(calendar_sync_title)
        calendar_row_layout.addSpacing(12)
        calendar_row_layout.addWidget(self.calendar_selection_label)
        calendar_row_layout.addStretch(1)
        calendar_row_layout.addWidget(self.calendar_selection_btn)
        cal_form.addRow(calendar_row)
        cal_form.addRow(hint)
        content_layout.addWidget(cal)

        account = QGroupBox("Cuenta de Google")
        acct_layout = QVBoxLayout(account)
        acct_hint = QLabel(
            "Elimina las credenciales guardadas en este equipo. La próxima vez que use "
            "calendario o correo deberá volver a iniciar sesión."
        )
        acct_hint.setObjectName("dialogHint")
        acct_hint.setWordWrap(True)
        acct_layout.addWidget(acct_hint)
        disconnect_btn = QPushButton("Desvincular cuenta de Google…")
        disconnect_btn.clicked.connect(self._on_disconnect_google)
        acct_layout.addWidget(disconnect_btn)
        content_layout.addWidget(account)

        content_layout.addStretch(1)

        btn_row = QHBoxLayout()
        reset_btn = QPushButton("Restablecer valores predeterminados…")
        reset_btn.clicked.connect(self._reset_to_defaults)
        btn_row.addWidget(reset_btn)
        btn_row.addStretch()
        cancel_btn = QPushButton("Cancelar")
        cancel_btn.clicked.connect(self.reject)
        save_btn = QPushButton("Guardar")
        save_btn.setObjectName("primaryButton")
        save_btn.clicked.connect(self._save)
        btn_row.addWidget(cancel_btn)
        btn_row.addWidget(save_btn)
        btn_row.setContentsMargins(20, 0, 20, 20)
        outer.addLayout(btn_row)

    def _open_template(self):
        dlg = TemplateEditorDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            mw = self.parent()
            if mw and hasattr(mw, "log_message"):
                mw.log_message("✓ Plantilla de email actualizada")

    def _on_disconnect_google(self):
        if not CREDENTIALS_PATH.exists():
            QMessageBox.information(
                self,
                "Sin sesión guardada",
                "No hay credenciales de Google guardadas en este equipo.",
            )
            return
        reply = QMessageBox.question(
            self,
            "Desvincular cuenta de Google",
            "Se revocará el acceso de esta aplicación en Google (si la red lo permite) y "
            "se borrarán las credenciales guardadas en este equipo.\n\n"
            "No podrá sincronizar el calendario ni enviar correos hasta volver a iniciar sesión.\n\n"
            "¿Desea continuar?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        revoke_and_remove_google_credentials()
        QMessageBox.information(
            self,
            "Cuenta desvinculada",
            "Las credenciales locales se eliminaron. Inicie sesión de nuevo cuando use "
            "sincronización o envío de correo.",
        )
        mw = self.parent()
        if mw and hasattr(mw, "log_message"):
            mw.log_message("✓ Cuenta de Google desvinculada en este equipo")
        if mw and hasattr(mw, "update_google_status_dot"):
            mw.update_google_status_dot()

    def _refresh_calendar_selection_label(self):
        calendar_id, calendar_name = get_selected_google_calendar()
        self.calendar_selection_label.setText(f"{calendar_name} ({calendar_id})")

    def _change_google_calendar(self):
        try:
            if not is_google_session_synced():
                mw = self.parent()
                service = get_google_service(
                    "calendar", embedded_oauth=True, parent=mw or self
                )
                if mw and hasattr(mw, "log_message"):
                    mw.log_message("✅ Inicio de sesión con Google completado correctamente.")
            else:
                service = get_google_service("calendar")
            selected = prompt_google_calendar_selection(self, service)
            if not selected:
                return
            self._refresh_calendar_selection_label()
            mw = self.parent()
            if mw and hasattr(mw, "log_message"):
                mw.log_message(
                    f"✓ Calendario de Google seleccionado: {selected['name']}"
                )
        except Exception as e:
            QMessageBox.warning(
                self,
                "Calendario de Google",
                f"No se pudo cambiar el calendario: {str(e)}",
            )

    def _reset_to_defaults(self):
        reply = QMessageBox.question(
            self,
            "Restablecer configuración",
            "Se aplicarán los valores predeterminados:\n\n"
            "• Tema oscuro\n"
            "• Tamaño de visualización 100 %\n"
            "• Inicio con Windows activado\n"
            "• Notificaciones de PreCita activadas\n"
            "• Revisión cada 5 minutos\n"
            "• Sincronización de calendario: 15 días\n\n"
            "¿Continuar?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        set_setting("theme", DEFAULT_THEME)
        set_setting("display_scale_percent", str(DEFAULT_DISPLAY_SCALE_PERCENT))
        set_setting(
            "reminder_interval_sec", str(DEFAULT_REMINDER_INTERVAL_SEC)
        )
        set_setting("calendar_sync_days", str(DEFAULT_CALENDAR_SYNC_DAYS))
        set_setting("google_calendar_id", "primary")
        set_setting("google_calendar_name", "Calendario principal")
        set_setting("windows_notifications_enabled", DEFAULT_WINDOWS_NOTIFICATIONS)

        startup_ok = True
        startup_err = ""
        if sys.platform == "win32":
            startup_ok, startup_err = windows_startup_set(
                DEFAULT_WINDOWS_STARTUP == "1"
            )
            if startup_ok:
                set_setting("windows_startup", DEFAULT_WINDOWS_STARTUP)
        else:
            set_setting("windows_startup", DEFAULT_WINDOWS_STARTUP)

        if sys.platform == "win32" and not startup_ok:
            QMessageBox.warning(
                self,
                "Inicio con Windows",
                "No se pudo desactivar el inicio automático en el registro: "
                f"{startup_err}\n\n"
                "El tema y el intervalo de revisión sí se restablecieron.",
            )

        idx_theme = self.theme_combo.findData(DEFAULT_THEME)
        if idx_theme >= 0:
            self.theme_combo.setCurrentIndex(idx_theme)
        idx_scale = self.display_scale_combo.findData(DEFAULT_DISPLAY_SCALE_PERCENT)
        self.display_scale_combo.setCurrentIndex(idx_scale if idx_scale >= 0 else 0)
        self.startup_check.setChecked(DEFAULT_WINDOWS_STARTUP == "1")
        self.notifications_check.setChecked(DEFAULT_WINDOWS_NOTIFICATIONS == "1")
        ix_iv = self.interval_combo.findData(DEFAULT_REMINDER_INTERVAL_SEC)
        self.interval_combo.setCurrentIndex(ix_iv if ix_iv >= 0 else 0)
        self.sync_days_input.setText(str(DEFAULT_CALENDAR_SYNC_DAYS))
        self._refresh_calendar_selection_label()

        app = QApplication.instance()
        if app:
            apply_app_appearance(app, DEFAULT_THEME, DEFAULT_DISPLAY_SCALE_PERCENT)

        mw = self.parent()
        if mw and hasattr(mw, "reminder_timer"):
            mw.reminder_timer.stop()
            mw.reminder_timer.start(DEFAULT_REMINDER_INTERVAL_SEC * 1000)

        if mw and hasattr(mw, "log_message"):
            mw.log_message("✓ Configuración restablecida a valores predeterminados")

    def _save(self):
        theme = self.theme_combo.currentData()
        scale_percent = int(self.display_scale_combo.currentData())
        set_setting("theme", theme)
        set_setting("display_scale_percent", str(scale_percent))
        app = QApplication.instance()
        if app:
            apply_app_appearance(app, theme, scale_percent)

        mw = self.parent()

        sec = int(self.interval_combo.currentData())
        set_setting("reminder_interval_sec", str(sec))
        raw_days = self.sync_days_input.text().strip()
        if not raw_days.isdigit():
            QMessageBox.warning(
                self,
                "Días de sincronización inválidos",
                "Ingrese solo números para los días de sincronización del calendario.",
            )
            return
        sync_days = int(raw_days)
        if not (MIN_CALENDAR_SYNC_DAYS <= sync_days <= MAX_CALENDAR_SYNC_DAYS):
            QMessageBox.warning(
                self,
                "Días de sincronización fuera de rango",
                f"Debe ingresar un valor entre {MIN_CALENDAR_SYNC_DAYS} y "
                f"{MAX_CALENDAR_SYNC_DAYS}.",
            )
            return
        set_setting("calendar_sync_days", str(sync_days))
        set_setting(
            "windows_notifications_enabled",
            "1" if self.notifications_check.isChecked() else "0",
        )
        if mw and hasattr(mw, "reminder_timer"):
            mw.reminder_timer.stop()
            mw.reminder_timer.start(sec * 1000)

        if sys.platform == "win32":
            want = self.startup_check.isChecked()
            ok, err = windows_startup_set(want)
            if ok:
                set_setting("windows_startup", "1" if want else "0")
            else:
                QMessageBox.warning(
                    self,
                    "Inicio con Windows",
                    f"No se pudo actualizar el registro: {err}",
                )

        if mw and hasattr(mw, "log_message"):
            mw.log_message("✓ Configuración guardada")

        self.accept()


class HelpDialog(QDialog):
    """Ventana de ayuda e información general de PreCita."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Ayuda — PreCita — {get_version()}")
        self.setMinimumWidth(700)
        self.resize(760, 560)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(18, 18, 18, 18)
        outer.setSpacing(10)

        blocked_exts = ", ".join(sorted(BLOCKED_ATTACHMENT_EXTENSIONS))
        archive_exts = ", ".join(sorted([x for x in ARCHIVE_EXTENSIONS if x != ".zip"]))
        help_data_dir_url = QUrl.fromLocalFile(str(PRECITA_LP)).toString()

        body = QTextBrowser()
        body.setReadOnly(True)
        body.setOpenExternalLinks(False)
        body.setOpenLinks(False)
        body.anchorClicked.connect(self._handle_help_link)
        body.setStyleSheet("font-size: 14px;")
        body.setHtml(
            f"""
            <h2>PreCita</h2>
            <p>
                PreCita es una aplicación de escritorio Windows para organizar citas, mantener
                contactos y automatizar recordatorios por correo electrónico con Google Calendar y Gmail.
                Los datos se conservan localmente en su equipo. Está completamente desarrollada en Python 
                3.12 por <code>eucarigo</code>, programador independiente (consulte 
                <a href="https://eucarigo.com/">https://eucarigo.com/</a>).
            </p>

            <h3 style="margin-top: 32px;">Características principales</h3>
            <ul>
                <li>Vistas de calendario diaria, semanal y mensual.</li>
                <li>Gestión de contactos y vinculación con citas.</li>
                <li>Plantilla de correo editable (asunto y cuerpo) con variables dinámicas.</li>
                <li>Envío de recordatorios manual (Lanzar pendientes) y automático por intervalo.</li>
                <li>Adjuntos reutilizables con reglas de seguridad para evitar bloqueos en Gmail.</li>
                <li>Vista de <a href="app://open_storage">Datos locales</a> para revisar, optimizar y encriptar la base de datos local.</li>
                <li>Indicador visual de sesión de Google y cambio de calendario sincronizado.</li>
            </ul>

            <h3 style="margin-top: 32px;">Breve tutorial de uso</h3>
            <ol>
                <li>Pulse <a href="app://sync_calendar">Sincronizar</a> para iniciar sesión con Google y actualizar citas.</li>
                <li>Cree o revise contactos en <a href="app://open_new_contact">Anadir contacto</a> y <a href="app://open_contacts">Gestión de contactos</a>.</li>
                <li>Abra <a href="app://open_template_editor">Personalizar plantilla</a> y ajuste asunto, cuerpo y adjuntos.</li>
                <li>Revise <a href="app://open_settings">Configuracion</a> para tema, intervalo, días de sincronización y calendario de Google.</li>
                <li>Entre en <a href="app://open_storage">Datos locales</a> para comprobar el estado de <code>~\\.precita\\</code>, usar <code>Optimizar</code> (VACUUM) y gestionar <code>Encriptar base de datos</code>.</li>
                <li>Ejecute <a href="app://send_reminders">Lanzar pendientes</a> para envío manual o deje activo el envío automático.</li>
            </ol>

            <h3 style="margin-top: 32px;">Variables de plantilla</h3>
            <p>
                Puede usar las siguientes variables en el asunto y cuerpo del correo:
            </p>
            <p>
                <code>{{nombre_citado}}</code>, <code>{{apellidos_citado}}</code>, <code>{{correo_citado}}</code>,
                <code>{{tlf_citado}}</code>, <code>{{hora_cita}}</code>, <code>{{fecha_cita}}</code> y
                <code>{{dia_semana}}</code>.
            </p>

            <h3 style="margin-top: 32px;">Adjuntos y seguridad</h3>
            <p>
                Para cumplir la política de Gmail y reducir errores de envío, PreCita bloquea archivos 
                de riesgo al cargarlos como adjuntos en la plantilla personalizable. Más información:
                <a href="https://support.google.com/mail/answer/6584">https://support.google.com/mail/answer/6584</a>.
            </p>

            <p>Extensiones bloqueadas:</p>
            <p><code>{blocked_exts}</code></p>
            <p>Acerca de las carpetas comprimidas:</p>
            <ul>
                <li>Se permite <code>.zip</code>, pero se bloquea si contiene archivos no permitidos.</li>
                <li>No se permiten otros comprimidos: <code>{archive_exts}</code>.</li>
                <li>El tamaño total de cuerpo + adjuntos esta limitado a 16 MB.</li>
            </ul>

            <h3 style="margin-top: 32px;">Tratamiento de datos</h3>
            <ul>
                <li>Los datos de contactos, citas y configuración se guardan localmente en su equipo en
                    <a href="{help_data_dir_url}"><code>~\\.precita\\</code></a>.
                </li>
                <li>Las credenciales de Google se almacenan localmente para permitir sincronización y envío.</li>
                <li>Sólo se comparte información con las APIs de Google necesarias para la funcionalidad elegida.</li>
                <li>Puede revocar y eliminar credenciales desde <a href="app://open_settings">Configuración</a>.</li>
                <li>Puede revisar tamaños, archivos clave y ejecutar mantenimiento desde <a href="app://open_storage">Datos locales</a>.</li>
                <li>Desde <a href="app://open_storage">Datos locales</a> puede abrir <code>Encriptar base de datos</code> para habilitar/deshabilitar el cifrado de <code>precita.db</code>.</li>
                <li>Si la encriptación está habilitada, PreCita solicitará la clave al iniciar para desbloquear la base de datos.</li>
                <li>La clave se aplica directamente sobre el cifrado del fichero local para facilitar su uso desde herramientas externas autorizadas.</li>
                <li>No se realiza telemetría ni analítica de ningún tipo.</li>
            </ul>
            <p>
                Puede entrar en el repositorio de GitHub oficial de PreCita para comprobar su ingeniería con total 
                trasparencia: <a href="https://github.com/eucarigo/precita/">https://github.com/eucarigo/precita/</a>.
            </p>

            <h3 style="margin-top: 32px;">Licencia y soporte técnico</h3>
            <p>
                PreCita es software distribuido bajo GNU GPLU v3, sin garantías. Puede visitar 
                <a href="https://www.gnu.org/licenses/gpl-3.0.txt">https://www.gnu.org/licenses/gpl-3.0.txt</a>
                para más información acerca de esta licencia.<br><br>Para más ayuda, puede visitar la página oficial
                del desarrollador <a href="https://eucarigo.com/precita/">https://eucarigo.com/precita</a> o 
                enviar un correo electrónico a <a href="mailto:help@eucarigo.com?subject=PreCita">help@eucarigo.com</a> 
                con los detalles de su problema, sugerencia o duda.
            </p>
            """
        )
        outer.addWidget(body)

        close_row = QHBoxLayout()
        close_row.addStretch()
        close_btn = QPushButton("Cerrar")
        close_btn.clicked.connect(self.accept)
        close_row.addWidget(close_btn)
        outer.addLayout(close_row)

    def _handle_help_link(self, url):
        """Gestionar enlaces internos y externos desde Ayuda."""
        if url.scheme() != "app":
            QDesktopServices.openUrl(url)
            return

        parent_window = self.parent()
        if not parent_window:
            return

        route = url.host() or url.path().lstrip("/")
        if route == "open_contacts" and hasattr(parent_window, "open_contacts_dialog"):
            self.accept()
            parent_window.open_contacts_dialog()
        elif route == "open_new_contact" and hasattr(parent_window, "open_new_contact_dialog"):
            self.accept()
            parent_window.open_new_contact_dialog()
        elif route == "open_template_editor" and hasattr(parent_window, "open_template_editor"):
            self.accept()
            parent_window.open_template_editor()
        elif route == "open_settings" and hasattr(parent_window, "open_app_settings"):
            self.accept()
            parent_window.open_app_settings()
        elif route == "open_storage" and hasattr(parent_window, "open_storage_dialog"):
            self.accept()
            parent_window.open_storage_dialog()
        elif route == "sync_calendar" and hasattr(parent_window, "sync_calendar"):
            self.accept()
            parent_window.sync_calendar()
        elif route == "send_reminders" and hasattr(parent_window, "send_reminders"):
            self.accept()
            parent_window.send_reminders()


class NoPastePasswordLineEdit(QLineEdit):
    """Campo de contraseña que bloquea el pegado desde portapapeles."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.NoContextMenu)

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Paste):
            event.ignore()
            return
        super().keyPressEvent(event)


class DbEncryptionDialog(QDialog):
    """Dialogo para habilitar/deshabilitar la encriptacion del fichero SQLite."""

    OPTION_DISABLED = "Deshabilitar encriptación (predeterminado)"
    OPTION_ENABLED = "Habilitar encriptación (recomendable)"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Encriptar base de datos")
        self.setModal(True)
        self.setMinimumWidth(540)
        self.resize(580, 310)

        cfg = _load_db_encryption_config()
        self.previous_enabled = bool(cfg.get("enabled", False))

        outer = QVBoxLayout(self)
        outer.setContentsMargins(18, 18, 18, 18)
        outer.setSpacing(10)

        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setSpacing(10)

        self.mode_combo = QComboBox()
        self.mode_combo.addItem(self.OPTION_DISABLED, False)
        self.mode_combo.addItem(self.OPTION_ENABLED, True)
        self.mode_combo.setCurrentIndex(1 if self.previous_enabled else 0)
        form.addRow("Seleccione una opción:", self.mode_combo)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setPlaceholderText("sólo caracteres alfanuméricos")
        password_row_widget = QWidget()
        password_row_layout = QHBoxLayout(password_row_widget)
        password_row_layout.setContentsMargins(0, 0, 0, 0)
        password_row_layout.setSpacing(6)
        password_row_layout.addWidget(self.password_input, 1)
        self.password_toggle_btn = QToolButton()
        self.password_toggle_btn.setText("*")
        self.password_toggle_btn.setToolTip("Mostrar/Ocultar contraseña")
        self.password_toggle_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.password_toggle_btn.clicked.connect(
            lambda: self._toggle_password_visibility(self.password_input, self.password_toggle_btn)
        )
        password_row_layout.addWidget(self.password_toggle_btn)
        self.password_label = QLabel("Introduzca su contraseña:")
        form.addRow(self.password_label, password_row_widget)

        self.repeat_password_input = NoPastePasswordLineEdit()
        self.repeat_password_input.setEchoMode(QLineEdit.EchoMode.Password)
        repeat_row_widget = QWidget()
        repeat_row_layout = QHBoxLayout(repeat_row_widget)
        repeat_row_layout.setContentsMargins(0, 0, 0, 0)
        repeat_row_layout.setSpacing(6)
        repeat_row_layout.addWidget(self.repeat_password_input, 1)
        self.repeat_password_toggle_btn = QToolButton()
        self.repeat_password_toggle_btn.setText("*")
        self.repeat_password_toggle_btn.setToolTip("Mostrar/Ocultar contraseña")
        self.repeat_password_toggle_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.repeat_password_toggle_btn.clicked.connect(
            lambda: self._toggle_password_visibility(self.repeat_password_input, self.repeat_password_toggle_btn)
        )
        repeat_row_layout.addWidget(self.repeat_password_toggle_btn)
        self.repeat_password_label = QLabel("Repita su contraseña:")
        form.addRow(self.repeat_password_label, repeat_row_widget)

        hint_label = QLabel(
            "La clave introducida se usa directamente para encriptar/desencriptar la base de datos. "
            "Por favor, asegúrate de guardar en un lugar seguro la contraseña elegida. PreCita ni almacena "
            "ni se hace cargo de pérdidas de claves critpográficas escogidas por el usuario."
        )
        hint_label.setWordWrap(True)
        hint_label.setObjectName("dialogHint")

        outer.addLayout(form)
        outer.addWidget(hint_label)

        actions = QHBoxLayout()
        actions.addStretch()

        cancel_btn = QPushButton("Cancelar")
        cancel_btn.clicked.connect(self._attempt_cancel)
        actions.addWidget(cancel_btn)

        save_btn = QPushButton("Guardar cambios")
        save_btn.setObjectName("primaryButton")
        save_btn.clicked.connect(self._save_changes)
        actions.addWidget(save_btn)
        outer.addLayout(actions)
        self.mode_combo.currentIndexChanged.connect(self._sync_password_fields_state)
        self._sync_password_fields_state()

    def _current_enabled(self):
        return bool(self.mode_combo.currentData())

    def _toggle_password_visibility(self, line_edit, toggle_btn):
        currently_hidden = line_edit.echoMode() == QLineEdit.EchoMode.Password
        line_edit.setEchoMode(QLineEdit.EchoMode.Normal if currently_hidden else QLineEdit.EchoMode.Password)
        toggle_btn.setText("¡!" if currently_hidden else "*")

    def _sync_password_fields_state(self):
        enabled = self._current_enabled()
        self.password_label.setEnabled(enabled)
        self.repeat_password_label.setEnabled(enabled)
        self.password_input.setEnabled(enabled)
        self.repeat_password_input.setEnabled(enabled)
        self.password_toggle_btn.setEnabled(enabled)
        self.repeat_password_toggle_btn.setEnabled(enabled)

    def _has_unsaved_changes(self):
        return self._current_enabled() != self.previous_enabled

    def _validate_password_fields(self):
        password = (self.password_input.text() or "").strip()
        repeated = (self.repeat_password_input.text() or "").strip()

        if not password or not repeated:
            QMessageBox.warning(self, "Contraseña obligatoria", "Debe rellenar ambos campos de contraseña.")
            return None
        if not re.fullmatch(r"[A-Za-z0-9]+", password):
            QMessageBox.warning(
                self,
                "Contraseña no válida",
                "La contraseña sólo puede incluir caracteres alfanuméricos (sin espacios ni símbolos).",
            )
            return None
        if password != repeated:
            QMessageBox.warning(
                self,
                "Contraseñas distintas",
                "La contraseña y su repetición no coinciden.",
            )
            return None
        return password

    def _attempt_cancel(self):
        if self._has_unsaved_changes():
            decision = QMessageBox.question(
                self,
                "Cambios sin guardar",
                "Ha realizado cambios que no se han guardado. ¿Desea cerrar y descartarlos?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if decision != QMessageBox.StandardButton.Yes:
                return False
        super().reject()
        return True

    def reject(self):
        self._attempt_cancel()

    def _save_changes(self):
        global RUNTIME_DB_ENCRYPTION_ENABLED, RUNTIME_DB_ENCRYPTION_PASSWORD

        selected_enabled = self._current_enabled()
        password = None
        if selected_enabled:
            password = self._validate_password_fields()
            if password is None:
                return
        if selected_enabled and _is_db_file_encrypted(DB_PATH):
            try:
                decrypt_database_file(password)
            except (ValueError, OSError) as exc:
                QMessageBox.warning(
                    self,
                    "No se pudo validar la contraseña",
                    f"La contraseña no permite abrir la base de datos cifrada actual: {exc}",
                )
                return

        _save_db_encryption_config(selected_enabled)
        RUNTIME_DB_ENCRYPTION_ENABLED = selected_enabled
        if selected_enabled:
            RUNTIME_DB_ENCRYPTION_PASSWORD = password
        else:
            RUNTIME_DB_ENCRYPTION_PASSWORD = None
            if _is_db_file_encrypted(DB_PATH):
                password, ok = QInputDialog.getText(
                    self,
                    "Deshabilitar encriptación",
                    "Introduzca la contraseña actual para desencriptar la base de datos:",
                    QLineEdit.EchoMode.Password,
                )
                if not ok or not (password or "").strip():
                    QMessageBox.warning(
                        self,
                        "Contraseña obligatoria",
                        "Debe introducir la contraseña actual para deshabilitar la encriptación.",
                    )
                    _save_db_encryption_config(True)
                    RUNTIME_DB_ENCRYPTION_ENABLED = True
                    return
                try:
                    decrypt_database_file(password.strip())
                except (ValueError, OSError) as exc:
                    QMessageBox.warning(
                        self,
                        "No se pudo deshabilitar la encriptación",
                        f"No se pudo desencriptar la base de datos con esa clave: {exc}",
                    )
                    _save_db_encryption_config(True)
                    RUNTIME_DB_ENCRYPTION_ENABLED = True
                    return

        self.previous_enabled = selected_enabled
        QMessageBox.information(
            self,
            "Configuración guardada",
            "Se guardó la configuración de encriptación de la base de datos.",
        )
        self.accept()


class StorageDialog(QDialog):
    """Ventana para consultar el almacenamiento interno de PreCita."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Datos locales — PreCita — {get_version()}")
        self.setMinimumWidth(700)
        self.resize(760, 560)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(18, 18, 18, 18)
        outer.setSpacing(10)

        self.storage_body = QTextBrowser()
        self.storage_body.setReadOnly(True)
        self.storage_body.setOpenExternalLinks(False)
        self.storage_body.setOpenLinks(False)
        self.storage_body.setStyleSheet("font-size: 14px;")
        self.storage_body.setHtml(self._build_storage_html())
        outer.addWidget(self.storage_body)

        close_row = QHBoxLayout()
        encryption_btn = QPushButton("Encriptar base de datos")
        encryption_btn.setToolTip("Configura el cifrado de seguridad del fichero precita.db.")
        encryption_btn.clicked.connect(self._open_db_encryption_dialog)
        close_row.addWidget(encryption_btn)
        optimize_btn = QPushButton("Optimizar")
        optimize_btn.setToolTip("Ejecuta VACUUM para compactar la base de datos local.")
        optimize_btn.clicked.connect(self._optimize_storage)
        close_row.addWidget(optimize_btn)
        close_row.addStretch()
        close_btn = QPushButton("Cerrar")
        close_btn.clicked.connect(self.accept)
        close_row.addWidget(close_btn)
        outer.addLayout(close_row)

    def _format_bytes(self, amount):
        size = float(max(0, amount))
        units = ("B", "KB", "MB", "GB", "TB")
        unit_idx = 0
        while size >= 1024 and unit_idx < len(units) - 1:
            size /= 1024
            unit_idx += 1
        if unit_idx == 0:
            return f"{int(size)} {units[unit_idx]}"
        return f"{size:.2f} {units[unit_idx]}"

    def _scan_precita_storage(self):
        base_dir = PRECITA_LP
        total_size = 0
        files_count = 0
        dirs_count = 0
        if base_dir.exists():
            for root, dirnames, filenames in os.walk(base_dir):
                dirs_count += len(dirnames)
                files_count += len(filenames)
                for file_name in filenames:
                    file_path = Path(root) / file_name
                    try:
                        total_size += file_path.stat().st_size
                    except OSError:
                        continue
        return {
            "base_dir": base_dir,
            "total_size": total_size,
            "files_count": files_count,
            "dirs_count": dirs_count,
        }

    def _file_line(self, label, path):
        if not path.exists():
            return f"<li><code>{label}</code>: no encontrado</li>"
        try:
            stat_obj = path.stat()
            modified = datetime.fromtimestamp(stat_obj.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            return (
                f"<li><code>{label}</code>: "
                f"{self._format_bytes(stat_obj.st_size)} "
                f"(modificado: {modified})</li>"
            )
        except OSError:
            return f"<li><code>{label}</code>: no accesible</li>"

    def _build_storage_html(self):
        storage = self._scan_precita_storage()
        base_dir = storage["base_dir"]
        base_dir_url = QUrl.fromLocalFile(str(base_dir)).toString()
        attachments_count = 0
        if ATTACHMENTS_DIR.exists():
            try:
                attachments_count = sum(1 for p in ATTACHMENTS_DIR.iterdir() if p.is_file())
            except OSError:
                attachments_count = 0

        return f"""
            <h2>Datos locales de PreCita</h2>
            <p>
                Esta vista muestra información relevante del directorio local de datos:
                <a href="{base_dir_url}"><code>~\\.precita\\</code></a>.
            </p>

            <h3 style="margin-top: 28px;">Resumen general</h3>
            <ul>
                <li>Ruta real: <code>{base_dir}</code></li>
                <li>Elementos totales: <code>{storage["files_count"]}</code> archivo(s) y <code>{storage["dirs_count"]}</code> carpeta(s)</li>
                <li>Tamaño total aproximado: <code>{self._format_bytes(storage["total_size"])}</code></li>
                <li>Archivos de adjuntos en plantilla: <code>{attachments_count}</code></li>
            </ul>

            <h3 style="margin-top: 28px;">Archivos principales</h3>
            <ul>
                {self._file_line("precita.db", DB_PATH)}
                {self._file_line("token.json", CREDENTIALS_PATH)}
                {self._file_line("client_secret.json", CLIENT_SECRETS)}
            </ul>

            <h3 style="margin-top: 28px;">Carpeta de adjuntos</h3>
            <ul>
                <li>Ruta: <code>{ATTACHMENTS_DIR}</code></li>
                <li>Estado: <code>{"disponible" if ATTACHMENTS_DIR.exists() else "no encontrada"}</code></li>
            </ul>

            <h3 style="margin-top: 28px;">Botón "Optimizar"</h3>
            <p>
                El botón <code>Optimizar</code> ejecuta el comando <code>VACUUM</code> sobre el fichero 
                <code>precita.db</code>. Este proceso reconstruye y compacta la base de datos para
                recuperar espacio no utilizado y dejarla más ordenada internamente.
            </p>
            <ul>
                <li>No elimina citas, contactos ni configuraciones.</li>
                <li>Puede tardar más cuando la base de datos es grande.</li>
                <li>Al terminar, esta pantalla se actualiza automáticamente con los tamaños nuevos.</li>
            </ul>

            <h3 style="margin-top: 28px;">Botón "Encriptar base de datos"</h3>
            <p>
                El botón <code>Encriptar base de datos</code> abre una ventana para activar o desactivar
                la protección criptográfica de <code>precita.db</code> mediante una contraseña alfanumérica.
            </p>
            <ul>
                <li>Opciones disponibles: <code>Deshabilitar encriptación (predeterminado)</code> y <code>Habilitar encriptación (recomendable)</code>.</li>
                <li>Al habilitarla, PreCita pedirá la clave al abrir la aplicación para desbloquear la base de datos local.</li>
                <li>La contraseña es obligatoria al habilitar y debe coincidir en ambos campos.</li>
                <li>La clave elegida por el usuario se usa directamente para el cifrado local del fichero, sin transformación intermedia de formato.</li>
                <li>Si intenta cerrar la ventana con cambios pendientes, PreCita avisará antes de descartar.</li>
            </ul>
        """

    def _optimize_storage(self):
        if not DB_PATH.exists():
            QMessageBox.warning(
                self,
                "Optimización no disponible",
                "No se encontró la base de datos local en este equipo.",
            )
            return
        try:
            conn = sqlite3.connect(DB_PATH)
            conn.execute("VACUUM")
            conn.close()
        except sqlite3.Error as exc:
            QMessageBox.warning(
                self,
                "Error al optimizar",
                f"No se pudo ejecutar VACUUM: {exc}",
            )
            return
        self.storage_body.setHtml(self._build_storage_html())
        QMessageBox.information(
            self,
            "Optimización completada",
            "Se ejecutó VACUUM correctamente sobre la base de datos de PreCita.",
        )

    def _open_db_encryption_dialog(self):
        dialog = DbEncryptionDialog(self)
        dialog.exec()


# ============================================================================
# VENTANA PRINCIPAL
# ============================================================================

class PreCitaMainWindow(QMainWindow):
    """Ventana principal de PreCita."""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"PreCita — {get_version()}")
        self.setMinimumSize(960, 640)
        self.resize(1180, 720)
        
        init_database()
        
        central_widget = QWidget()
        central_widget.setObjectName("centralRoot")
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        header = QFrame()
        header.setObjectName("appHeader")
        header.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        header_row = QHBoxLayout(header)
        header_row.setContentsMargins(20, 14, 20, 14)
        header_row.setSpacing(16)
        
        brand_col = QVBoxLayout()
        brand_col.setSpacing(2)
        title_wrap = QWidget()
        title_wrap.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
        title_row = QGridLayout(title_wrap)
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.setHorizontalSpacing(0)
        title_row.setVerticalSpacing(0)
        title = QLabel("PreCita")
        title.setObjectName("brandTitle")
        self.google_status_dot_header = QLabel()
        self.google_status_dot_header.setObjectName("googleStatusDot")
        self.google_status_dot_header.setToolTip("No sincronizado con Google")
        self.google_status_dot_header.setContentsMargins(0, 0, 0, 0)
        title_row.addWidget(title, 0, 0, 1, 1, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        subtitle = QLabel("Menos ausencias, más eficiencia. Tu agenda en piloto automático.")
        subtitle.setObjectName("brandSubtitle")
        brand_col.addWidget(title_wrap, 0, Qt.AlignmentFlag.AlignLeft)
        brand_col.addWidget(subtitle)
        header_row.addLayout(brand_col, 0)
        header_row.addStretch(1)
        
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        menu_btn = QToolButton()
        menu_btn.setObjectName("headerMenuButton")
        menu_btn.setText("☰")
        menu_btn.setToolTip("Acciones")
        menu_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        menu_badge_wrap = QWidget()
        menu_badge_wrap.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        menu_badge_layout = QGridLayout(menu_badge_wrap)
        menu_badge_layout.setContentsMargins(0, 0, 0, 0)
        menu_badge_layout.setHorizontalSpacing(0)
        menu_badge_layout.setVerticalSpacing(0)
        menu_badge_layout.addWidget(menu_btn, 0, 0, 1, 1, Qt.AlignmentFlag.AlignCenter)
        menu_badge_layout.addWidget(
            self.google_status_dot_header,
            0,
            0,
            1,
            1,
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight,
        )

        actions_menu = QMenu(menu_btn)
        self.actions_menu = actions_menu
        sync_menu_widget = QWidget(actions_menu)
        self.sync_menu_widget = sync_menu_widget
        sync_menu_widget.setObjectName("syncMenuActionWidget")
        sync_menu_widget.setAttribute(Qt.WidgetAttribute.WA_Hover, True)
        sync_menu_layout = QHBoxLayout(sync_menu_widget)
        sync_menu_layout.setContentsMargins(12, 7, 12, 7)
        sync_menu_layout.setSpacing(6)
        self.sync_menu_button = QPushButton("Sincronizar")
        self.sync_menu_button.setObjectName("syncMenuActionButton")
        self.sync_menu_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.sync_menu_button.setFlat(True)
        self.sync_menu_button.clicked.connect(self.sync_calendar)
        self.sync_menu_button.clicked.connect(actions_menu.hide)
        self.google_status_dot_menu = QLabel()
        self.google_status_dot_menu.setObjectName("googleStatusDot")
        self.google_status_dot_menu.setToolTip("No sincronizado con Google")
        self.google_status_dot_menu.setCursor(Qt.CursorShape.PointingHandCursor)
        self.sync_menu_widget.installEventFilter(self)
        self.google_status_dot_menu.installEventFilter(self)
        sync_menu_layout.addWidget(self.sync_menu_button)
        sync_menu_layout.addWidget(self.google_status_dot_menu, 0, Qt.AlignmentFlag.AlignVCenter)
        sync_menu_layout.addStretch(1)
        sync_widget_action = QWidgetAction(actions_menu)
        sync_widget_action.setDefaultWidget(sync_menu_widget)
        actions_menu.addAction(sync_widget_action)
        self.launch_pending_menu_widget = QWidget(actions_menu)
        self.launch_pending_menu_widget.setObjectName("launchPendingMenuActionWidget")
        self.launch_pending_menu_widget.setAttribute(Qt.WidgetAttribute.WA_Hover, True)
        self.launch_pending_menu_widget.setToolTip("No sincronizado con Google")
        launch_pending_menu_layout = QHBoxLayout(self.launch_pending_menu_widget)
        launch_pending_menu_layout.setContentsMargins(12, 7, 12, 7)
        launch_pending_menu_layout.setSpacing(6)
        self.launch_pending_menu_button = QPushButton("Lanzar pendientes")
        self.launch_pending_menu_button.setObjectName("launchPendingMenuActionButton")
        self.launch_pending_menu_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.launch_pending_menu_button.setFlat(True)
        self.launch_pending_menu_button.setToolTip("No sincronizado con Google")
        self.launch_pending_menu_button.clicked.connect(self.send_reminders)
        self.launch_pending_menu_button.clicked.connect(actions_menu.hide)
        launch_pending_menu_layout.addWidget(self.launch_pending_menu_button)
        launch_pending_menu_layout.addStretch(1)
        launch_pending_widget_action = QWidgetAction(actions_menu)
        launch_pending_widget_action.setDefaultWidget(self.launch_pending_menu_widget)
        actions_menu.addAction(launch_pending_widget_action)
        actions_menu.addAction("Gestión de contactos", self.open_contacts_dialog)
        actions_menu.addAction("Datos locales", self.open_storage_dialog)
        actions_menu.addSeparator()
        actions_menu.addAction("Añadir contacto", self.open_new_contact_dialog)
        actions_menu.addAction("Personalizar plantilla", self.open_template_editor)
        actions_menu.addSeparator()
        check_updates_menu_widget = QWidget(actions_menu)
        check_updates_menu_widget.setObjectName("launchPendingMenuActionWidget")
        check_updates_menu_widget.setAttribute(Qt.WidgetAttribute.WA_Hover, True)
        check_updates_menu_widget.setToolTip("Esta funcionalidad será añadida en versiones futuras")
        check_updates_menu_widget.setCursor(Qt.CursorShape.ForbiddenCursor)
        check_updates_menu_layout = QHBoxLayout(check_updates_menu_widget)
        check_updates_menu_layout.setContentsMargins(12, 7, 12, 7)
        check_updates_menu_layout.setSpacing(6)
        check_updates_menu_button = QPushButton("Buscar actualizaciones")
        check_updates_menu_button.setObjectName("launchPendingMenuActionButton")
        check_updates_menu_button.setCursor(Qt.CursorShape.ForbiddenCursor)
        check_updates_menu_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        check_updates_menu_button.setFlat(True)
        check_updates_menu_button.setToolTip("Esta funcionalidad será añadida en versiones futuras")
        check_updates_menu_button.setEnabled(False)
        check_updates_menu_widget.setEnabled(False)
        check_updates_menu_layout.addWidget(check_updates_menu_button)
        check_updates_menu_layout.addStretch(1)
        check_updates_widget_action = QWidgetAction(actions_menu)
        check_updates_widget_action.setDefaultWidget(check_updates_menu_widget)
        actions_menu.addAction(check_updates_widget_action)
        actions_menu.addAction("Configuración", self.open_app_settings)
        actions_menu.addAction("Ayuda", self.open_help_dialog)
        actions_menu.addSeparator()
        self.logout_menu_widget = QWidget(actions_menu)
        self.logout_menu_widget.setObjectName("logoutMenuActionWidget")
        self.logout_menu_widget.setAttribute(Qt.WidgetAttribute.WA_Hover, True)
        self.logout_menu_widget.setToolTip("No sincronizado con Google")
        logout_menu_layout = QHBoxLayout(self.logout_menu_widget)
        logout_menu_layout.setContentsMargins(12, 7, 12, 7)
        logout_menu_layout.setSpacing(6)
        self.logout_menu_button = QPushButton("Cerrar sesión")
        self.logout_menu_button.setObjectName("logoutMenuActionButton")
        self.logout_menu_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.logout_menu_button.setFlat(True)
        self.logout_menu_button.setToolTip("No sincronizado con Google")
        self.logout_menu_button.clicked.connect(self.close_google_session)
        self.logout_menu_button.clicked.connect(actions_menu.hide)
        logout_menu_layout.addWidget(self.logout_menu_button)
        logout_menu_layout.addStretch(1)
        logout_widget_action = QWidgetAction(actions_menu)
        logout_widget_action.setDefaultWidget(self.logout_menu_widget)
        actions_menu.addAction(logout_widget_action)
        menu_btn.setMenu(actions_menu)
        btn_row.addWidget(menu_badge_wrap)
        header_row.addLayout(btn_row, 0)
        layout.addWidget(header)
        
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(16, 16, 16, 12)
        body_layout.setSpacing(14)

        self.current_view = "weekly"
        self.anchor_date = date.today()
        
        citas_card = QFrame()
        citas_card.setObjectName("panelCard")
        citas_inner = QVBoxLayout(citas_card)
        citas_inner.setContentsMargins(14, 12, 14, 12)
        citas_inner.setSpacing(8)
        citas_label = QLabel("Calendario")
        citas_label.setObjectName("sectionTitle")
        meta = QLabel("Consulte sus citas para los próximos días")
        meta.setObjectName("brandSubtitle")
        citas_inner.addWidget(citas_label)
        citas_inner.addWidget(meta)

        calendar_toolbar = QFrame()
        calendar_toolbar.setObjectName("calendarToolbar")
        toolbar_layout = QHBoxLayout(calendar_toolbar)
        toolbar_layout.setContentsMargins(10, 8, 10, 8)
        toolbar_layout.setSpacing(8)

        self.today_btn = QPushButton("Hoy")
        self.today_btn.setObjectName("calendarTodayButton")
        self.today_btn.clicked.connect(self.go_to_today)
        toolbar_layout.addWidget(self.today_btn)

        self.prev_btn = QPushButton("◀")
        self.prev_btn.setObjectName("calendarNavButton")
        self.prev_btn.clicked.connect(lambda: self.shift_period(-1))
        toolbar_layout.addWidget(self.prev_btn)

        self.next_btn = QPushButton("▶")
        self.next_btn.setObjectName("calendarNavButton")
        self.next_btn.clicked.connect(lambda: self.shift_period(1))
        toolbar_layout.addWidget(self.next_btn)

        self.period_title = QLabel("")
        self.period_title.setObjectName("calendarPeriodTitle")
        toolbar_layout.addWidget(self.period_title)
        toolbar_layout.addStretch(1)

        self.view_button_group = QButtonGroup(self)
        self.view_button_group.setExclusive(True)
        self.view_buttons = {}

        for label, view_key in (("Diario", "daily"), ("Semanal", "weekly"), ("Mensual", "monthly")):
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setObjectName("calendarViewButton")
            btn.clicked.connect(lambda checked, vk=view_key: self.set_calendar_view(vk))
            self.view_button_group.addButton(btn)
            self.view_buttons[view_key] = btn
            toolbar_layout.addWidget(btn)

        self.view_buttons[self.current_view].setChecked(True)
        citas_inner.addWidget(calendar_toolbar)
        
        self.appointments_table = QTableWidget()
        self.appointments_table.setColumnCount(6)
        self.appointments_table.setHorizontalHeaderLabels([
            "Cita", "Fecha", "Hora", "Contacto", "Estado", "Acción"
        ])
        self.appointments_table.setAlternatingRowColors(True)
        self.appointments_table.setShowGrid(False)
        self.appointments_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.appointments_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.appointments_table.verticalHeader().setVisible(False)
        self.appointments_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.appointments_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.appointments_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.appointments_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.appointments_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.appointments_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        self.appointments_table.setColumnWidth(5, 118)
        self.appointments_table.setMinimumHeight(180)
        self.appointments_table.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        citas_inner.addWidget(self.appointments_table)
        body_layout.addWidget(citas_card, 1)

        self.log_text = QTextEdit()
        self.log_text.setObjectName("activityLog")
        self.log_text.setReadOnly(True)
        self.log_text.setVisible(False)
        
        layout.addWidget(body, 1)
        
        # ===== TIMERS =====
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.refresh_tables)
        self.update_timer.start(30000)  # Cada 30 segundos
        
        self.reminder_timer = QTimer()
        self.reminder_timer.timeout.connect(self.send_reminders_auto)
        try:
            _riv = int(
                get_setting(
                    "reminder_interval_sec",
                    str(DEFAULT_REMINDER_INTERVAL_SEC),
                )
                or str(DEFAULT_REMINDER_INTERVAL_SEC)
            )
        except (TypeError, ValueError):
            _riv = DEFAULT_REMINDER_INTERVAL_SEC
        self.reminder_timer.start(max(60, _riv) * 1000)
        
        self.configure_appointments_table_for_view()
        self.update_period_title()
        self._zoom_shortcuts = []
        self._setup_zoom_shortcuts()
        
        self._help_shortcut = QShortcut(QKeySequence("Ctrl+H"), self)
        self._help_shortcut.setContext(Qt.ShortcutContext.ApplicationShortcut)
        self._help_shortcut.activated.connect(self.open_help_dialog)

        self._config_shortcut = QShortcut(QKeySequence("Ctrl+,"), self)
        self._config_shortcut.setContext(Qt.ShortcutContext.ApplicationShortcut)
        self._config_shortcut.activated.connect(self.open_app_settings)


        # Cargar datos iniciales
        self.refresh_tables()
        self.create_tray_icon()
        self.update_google_status_dot()
        self.log_message("✓ PreCita iniciado - Listo para usar con tu cuenta de Google")

    def _setup_zoom_shortcuts(self):
        shortcut_map = (
            ("Ctrl++", 1),
            ("Ctrl+=", 1),
            ("Ctrl+Plus", 1),
            ("Ctrl+-", -1),
            ("Ctrl+Minus", -1),
        )
        for sequence, step in shortcut_map:
            shortcut = QShortcut(QKeySequence(sequence), self)
            shortcut.setContext(Qt.ShortcutContext.ApplicationShortcut)
            shortcut.activated.connect(
                lambda step_value=step: self._adjust_display_scale(step_value)
            )
            self._zoom_shortcuts.append(shortcut)

    def _adjust_display_scale(self, step):
        if step == 0:
            return
        choices = list(DISPLAY_SCALE_CHOICES_PERCENT)
        current_scale = get_display_scale_percent()
        if current_scale in choices:
            current_index = choices.index(current_scale)
        else:
            current_index = choices.index(DEFAULT_DISPLAY_SCALE_PERCENT)
        target_index = max(0, min(len(choices) - 1, current_index + step))
        if target_index == current_index:
            return
        new_scale = choices[target_index]
        set_setting("display_scale_percent", str(new_scale))
        theme = get_setting("theme", DEFAULT_THEME) or DEFAULT_THEME
        app = QApplication.instance()
        if app:
            apply_app_appearance(app, theme, new_scale)
        self._show_display_scale_tooltip(new_scale)
        self.log_message(f"✓ Tamaño de visualización ajustado a {new_scale} %")

    def _show_display_scale_tooltip(self, scale_percent):
        center_global = self.mapToGlobal(self.rect().center())
        QToolTip.hideText()
        QToolTip.showText(center_global, f"{scale_percent}%", self, self.rect(), 1400)

    def eventFilter(self, obj, event):
        if (
            obj in {getattr(self, "sync_menu_widget", None), getattr(self, "google_status_dot_menu", None)}
            and event.type() == QEvent.Type.MouseButtonRelease
            and hasattr(event, "button")
            and event.button() == Qt.MouseButton.LeftButton
        ):
            self.sync_calendar()
            if hasattr(self, "actions_menu") and self.actions_menu is not None:
                self.actions_menu.hide()
            return True
        return super().eventFilter(obj, event)
    
    def log_message(self, message):
        """Agrega un mensaje al log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def sync_calendar(self):
        """Inicia la sincronización con Google Calendar en un thread separado."""
        if not is_google_session_synced():
            try:
                self.log_message("🔐 Abriendo inicio de sesión de Google en PreCita...")
                get_google_service('calendar', embedded_oauth=True, parent=self)
                self.log_message("✅ Inicio de sesión con Google completado correctamente.")
                self.update_google_status_dot()
            except Exception as e:
                self.show_error(f"No se pudo iniciar sesión con Google: {str(e)}")
                return
        try:
            service = get_google_service("calendar")
            calendars = list_google_calendars(service)
            selected_id, selected_name = get_selected_google_calendar()
            has_secondary_calendars = any(cal["id"] != "primary" for cal in calendars)
            if len(calendars) > 1 and selected_id == "primary" and has_secondary_calendars:
                reply = QMessageBox.question(
                    self,
                    "Seleccion de calendario recomendada",
                    "Su cuenta tiene varios calendarios. Para evitar sincronizar "
                    "eventos no deseados, puede elegir explícitamente cuál "
                    "calendario leerá PreCita.\n\n"
                    "¿Quiere seleccionarlo ahora?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes,
                )
                if reply == QMessageBox.StandardButton.Yes:
                    selected = prompt_google_calendar_selection(self, service)
                    if selected is None:
                        self.log_message(
                            "ℹ️ Sincronización cancelada: no se seleccionó calendario."
                        )
                        return
                    selected_id = selected["id"]
                    selected_name = selected["name"]
            self.log_message(f"📅 Calendario en uso: {selected_name} ({selected_id})")
        except Exception as e:
            self.show_error(f"No se pudo validar el calendario configurado: {str(e)}")
            return

        self.sync_thread = QThread()
        self.worker = SyncWorker()
        self.worker.moveToThread(self.sync_thread)
        
        self.sync_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.sync_thread.quit)
        self.worker.error.connect(self.show_error)
        self.worker.appointments_found.connect(self.on_appointments_found)
        
        self.sync_thread.start()
        self.log_message("🔄 Sincronizando con Google Calendar...")
    
    def send_reminders_auto(self):
        """Envía recordatorios automáticamente según el intervalo configurado."""
        self.send_reminders(auto=True)
    
    def send_reminders(self, auto=False):
        """Envía recordatorios de citas para mañana."""
        self.reminder_thread = QThread()
        self.reminder_worker = ReminderWorker(auto_mode=auto)
        self.reminder_worker.moveToThread(self.reminder_thread)
        
        self.reminder_thread.started.connect(self.reminder_worker.run)
        self.reminder_worker.finished.connect(self.reminder_thread.quit)
        self.reminder_worker.error.connect(self.show_error)
        self.reminder_worker.reminder_sent.connect(self.log_message)
        self.reminder_worker.auto_reminder_email_sent.connect(
            self.show_auto_email_notification
        )
        self.reminder_worker.finished.connect(self.refresh_tables)
        
        self.reminder_thread.start()
        self.log_message("📧 Procesando recordatorios...")

    def show_auto_email_notification(self, contact_name, appointment_time):
        """Muestra una notificación nativa en Windows tras un envío automático."""
        if sys.platform != "win32":
            return
        if get_setting(
            "windows_notifications_enabled", DEFAULT_WINDOWS_NOTIFICATIONS
        ) != "1":
            return
        if not hasattr(self, "tray_icon") or self.tray_icon is None:
            return
        title = "PreCita: correo automático enviado"
        message = (
            f"Contacto: {contact_name}\n"
            f"Hora prevista de la cita: {appointment_time}"
        )
        self.tray_icon.showMessage(
            title,
            message,
            QSystemTrayIcon.MessageIcon.Information,
            8000,
        )
    
    def on_appointments_found(self, appointments):
        """Manejador cuando se encuentran nuevas citas."""
        self.update_google_status_dot()
        if appointments:
            self.log_message(f"✓ Se sincronizaron {len(appointments)} citas nuevas")
        else:
            self.log_message("ℹ️ El calendario está actualizado")
        self.refresh_tables()
    
    def show_error(self, error_msg):
        """Muestra un error."""
        self.update_google_status_dot()
        self.log_message(f"❌ {error_msg}")
        QMessageBox.critical(self, "Error", error_msg)

    def update_google_status_dot(self):
        """Actualiza el indicador visual de sincronización de Google."""
        synced = is_google_session_synced()
        state = "synced" if synced else "unsynced"
        tooltip = "Sincronizado con Google" if synced else "No sincronizado con Google"
        status_dots = [
            getattr(self, "google_status_dot_header", None),
            getattr(self, "google_status_dot_menu", None),
        ]
        for dot in status_dots:
            if dot is None:
                continue
            dot.setProperty("syncState", state)
            dot.setToolTip(tooltip)
            dot.style().unpolish(dot)
            dot.style().polish(dot)
            dot.update()
        logout_button = getattr(self, "logout_menu_button", None)
        if logout_button is not None:
            logout_button.setEnabled(synced)
        logout_widget = getattr(self, "logout_menu_widget", None)
        if logout_widget is not None:
            logout_widget.setEnabled(synced)
        launch_pending_button = getattr(self, "launch_pending_menu_button", None)
        if launch_pending_button is not None:
            launch_pending_button.setEnabled(synced)
        launch_pending_widget = getattr(self, "launch_pending_menu_widget", None)
        if launch_pending_widget is not None:
            launch_pending_widget.setEnabled(synced)
    
    def refresh_tables(self):
        """Actualiza las citas."""
        self.load_appointments()

    def set_calendar_view(self, view_key):
        if view_key not in {"daily", "weekly", "monthly"}:
            return
        self.current_view = view_key
        self.configure_appointments_table_for_view()
        self.update_period_title()
        self.load_appointments()

    def go_to_today(self):
        self.anchor_date = date.today()
        self.update_period_title()
        self.load_appointments()

    def shift_period(self, step):
        if self.current_view == "daily":
            self.anchor_date = self.anchor_date + timedelta(days=step)
        elif self.current_view == "weekly":
            self.anchor_date = self.anchor_date + timedelta(days=7 * step)
        else:
            month_index = (self.anchor_date.month - 1) + step
            new_year = self.anchor_date.year + (month_index // 12)
            new_month = (month_index % 12) + 1
            self.anchor_date = date(new_year, new_month, 1)
        self.update_period_title()
        self.load_appointments()

    def _calendar_range(self):
        if self.current_view == "daily":
            start = self.anchor_date
            end = start + timedelta(days=1)
        elif self.current_view == "weekly":
            start = self.anchor_date - timedelta(days=self.anchor_date.weekday())
            end = start + timedelta(days=7)
        else:
            start = self.anchor_date.replace(day=1)
            if start.month == 12:
                end = date(start.year + 1, 1, 1)
            else:
                end = date(start.year, start.month + 1, 1)
        return start, end

    def update_period_title(self):
        start, end = self._calendar_range()
        if self.current_view == "daily":
            weekday_name = SPANISH_WEEKDAY_NAMES[start.weekday()]
            month_abbr = SPANISH_MONTH_ABBR[start.month - 1]
            period_text = f"{weekday_name}, {start.day:02d} {month_abbr} {start.year}"
        elif self.current_view == "weekly":
            start_month_abbr = SPANISH_MONTH_ABBR[start.month - 1]
            end_date = end - timedelta(days=1)
            end_month_abbr = SPANISH_MONTH_ABBR[end_date.month - 1]
            period_text = (
                f"{start.day:02d} {start_month_abbr} - "
                f"{end_date.day:02d} {end_month_abbr} {end_date.year}"
            )
        else:
            month_name = SPANISH_MONTH_NAMES[start.month - 1]
            period_text = f"{month_name} {start.year}"
        self.period_title.setText(period_text)

    def configure_appointments_table_for_view(self):
        if self.current_view == "daily":
            headers = ["Hora", "Cita", "Contacto", "Estado", "Acción"]
            stretch_cols = [1, 2]
            fixed_action_col = 4
        elif self.current_view == "weekly":
            headers = ["Día", "Fecha", "Hora", "Cita", "Contacto", "Estado", "Acción"]
            stretch_cols = [3, 4]
            fixed_action_col = 6
        else:
            headers = ["Fecha", "Hora", "Cita", "Contacto", "Estado", "Acción"]
            stretch_cols = [2, 3]
            fixed_action_col = 5

        self.appointments_table.setColumnCount(len(headers))
        self.appointments_table.setHorizontalHeaderLabels(headers)
        hdr = self.appointments_table.horizontalHeader()
        for col in range(len(headers)):
            if col in stretch_cols:
                hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.Stretch)
            elif col == fixed_action_col:
                hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.Fixed)
            else:
                hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        self.appointments_table.setColumnWidth(fixed_action_col, 118)
    
    def load_appointments(self):
        """Carga las citas en la tabla."""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, event_title, event_date, contact_id, reminder_sent, email_sent_at
            FROM appointments
            ORDER BY event_date ASC
        ''')
        all_appointments = cursor.fetchall()
        conn.close()

        start_date, end_date = self._calendar_range()
        appointments = []
        for apt in all_appointments:
            apt_id, title, date_str, contact_id, sent, email_sent_at = apt
            # Parsear fecha
            try:
                apt_date = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                apt_local_date = apt_date.date()
            except Exception:
                apt_date = None
                apt_local_date = None

            if apt_local_date is None or not (start_date <= apt_local_date < end_date):
                continue

            date_formatted = apt_date.strftime("%d/%m/%Y")
            time_formatted = apt_date.strftime("%H:%M")
            day_name = SPANISH_WEEKDAY_NAMES[apt_date.weekday()]

            # Nombre del contacto
            contact_name = "-"
            if contact_id:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                cursor.execute(
                    'SELECT first_name, last_name FROM contacts WHERE id = ?',
                    (contact_id,),
                )
                result = cursor.fetchone()
                if result:
                    fn, ln = result
                    contact_name = contact_full_name(fn, ln) or "-"
                else:
                    contact_name = "-"
                conn.close()

            status = "Correo enviado" if sent else "Pendiente"
            status_color = QColor("#0d9488") if sent else QColor("#c2410c")

            appointments.append(
                {
                    "id": apt_id,
                    "title": title,
                    "date": date_formatted,
                    "time": time_formatted,
                    "day": day_name,
                    "contact": contact_name,
                    "status": status,
                    "status_color": status_color,
                }
            )

        self.appointments_table.setRowCount(len(appointments))

        for row, appointment in enumerate(appointments):
            if self.current_view == "daily":
                self.appointments_table.setItem(row, 0, QTableWidgetItem(appointment["time"]))
                self.appointments_table.setItem(row, 1, QTableWidgetItem(appointment["title"]))
                self.appointments_table.setItem(row, 2, QTableWidgetItem(appointment["contact"]))
                status_col = 3
                action_col = 4
            elif self.current_view == "weekly":
                self.appointments_table.setItem(row, 0, QTableWidgetItem(appointment["day"]))
                self.appointments_table.setItem(row, 1, QTableWidgetItem(appointment["date"]))
                self.appointments_table.setItem(row, 2, QTableWidgetItem(appointment["time"]))
                self.appointments_table.setItem(row, 3, QTableWidgetItem(appointment["title"]))
                self.appointments_table.setItem(row, 4, QTableWidgetItem(appointment["contact"]))
                status_col = 5
                action_col = 6
            else:
                self.appointments_table.setItem(row, 0, QTableWidgetItem(appointment["date"]))
                self.appointments_table.setItem(row, 1, QTableWidgetItem(appointment["time"]))
                self.appointments_table.setItem(row, 2, QTableWidgetItem(appointment["title"]))
                self.appointments_table.setItem(row, 3, QTableWidgetItem(appointment["contact"]))
                status_col = 4
                action_col = 5

            status_item = QTableWidgetItem(appointment["status"])
            status_item.setForeground(appointment["status_color"])
            self.appointments_table.setItem(row, status_col, status_item)

            # Botón de enviar email manual
            send_btn = QPushButton("Enviar ahora")
            send_btn.clicked.connect(
                lambda checked, aid=appointment["id"]: self.send_single_reminder(aid)
            )
            self.appointments_table.setCellWidget(row, action_col, send_btn)
    
    def send_single_reminder(self, appointment_id):
        """Envía recordatorio para una cita específica."""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT a.event_title, a.event_date, c.first_name, c.last_name, c.email, c.phone
            FROM appointments a
            LEFT JOIN contacts c ON a.contact_id = c.id
            WHERE a.id = ?
        ''', (appointment_id,))
        result = cursor.fetchone()
        conn.close()
        
        if not result:
            QMessageBox.warning(self, "Error", "Cita no encontrada")
            return
        
        title, event_date, c_first, c_last, contact_email, contact_phone = result
        template_vars = _build_email_template_variables(
            c_first,
            c_last,
            contact_email,
            contact_phone,
            event_date,
        )
        contact_display_name = contact_full_name(c_first, c_last) or "Paciente"
        
        if not contact_email:
            QMessageBox.warning(self, "Error", "No hay email asociado a esta cita")
            return
        
        success, message = send_reminder_email_gmail(
            appointment_id,
            contact_email,
            template_vars,
        )
        
        if success:
            self.log_message(f"✓ Email enviado a {contact_display_name}")
            QMessageBox.information(self, "Éxito", f"Email enviado a {contact_display_name}")
            self.refresh_tables()
        else:
            self.log_message(f"✗ Error: {message}")
            QMessageBox.critical(self, "Error", message)
    
    def load_contacts(self):
        """Carga los contactos en la tabla."""
        hdr = self.contacts_table.horizontalHeader()
        sort_col = hdr.sortIndicatorSection()
        sort_order = hdr.sortIndicatorOrder()
        if sort_col < 0:
            sort_col = 1
            sort_order = Qt.SortOrder.AscendingOrder

        self.contacts_table.setSortingEnabled(False)

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT id, first_name, last_name, email, phone FROM contacts
            ORDER BY last_name COLLATE NOCASE ASC, first_name COLLATE NOCASE ASC
            """
        )
        contacts = cursor.fetchall()
        conn.close()

        self.contacts_table.setRowCount(len(contacts))

        for row, (cid, first_name, last_name, email, phone) in enumerate(contacts):
            name_item = _CaseInsensitiveTableWidgetItem(first_name)
            name_item.setData(Qt.ItemDataRole.UserRole, cid)
            self.contacts_table.setItem(row, 0, name_item)
            self.contacts_table.setItem(
                row, 1, _CaseInsensitiveTableWidgetItem(last_name or "")
            )
            self.contacts_table.setItem(row, 2, _CaseInsensitiveTableWidgetItem(email))
            self.contacts_table.setItem(
                row, 3, _CaseInsensitiveTableWidgetItem(phone or "-")
            )

        self.contacts_table.setSortingEnabled(True)
        self.contacts_table.sortByColumn(sort_col, sort_order)
    
    def _contacts_context_menu(self, pos):
        """Menú contextual sobre una fila de contacto."""
        row = self.contacts_table.rowAt(pos.y())
        if row < 0:
            return
        name_item = self.contacts_table.item(row, 0)
        if not name_item:
            return
        cid = name_item.data(Qt.ItemDataRole.UserRole)
        if cid is None:
            return
        self.contacts_table.selectRow(row)
        cid = int(cid)
        menu = QMenu(self)
        menu.addAction("Editar nombre del contacto", lambda: self._edit_contact_first_name(cid))
        menu.addAction(
            "Editar apellido(s) del contacto", lambda: self._edit_contact_last_name(cid)
        )
        menu.addAction("Editar correo del contacto", lambda: self._edit_contact_email(cid))
        menu.addAction("Editar teléfono del contacto", lambda: self._edit_contact_phone(cid))
        menu.addSeparator()
        delete_action = QAction("Eliminar contacto", menu)
        delete_action.setObjectName("precitaMenuDelete")
        delete_action.setProperty("precitaDestructive", True)
        delete_action.triggered.connect(lambda: self._delete_contact(cid))
        menu.addAction(delete_action)
        menu.exec(self.contacts_table.viewport().mapToGlobal(pos))
    
    def _edit_contact_first_name(self, contact_id):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT first_name FROM contacts WHERE id = ?', (contact_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return
        current = row[0]
        text, ok = QInputDialog.getText(
            self, "Editar nombre", "Nombre del contacto:", text=current
        )
        if not ok:
            return
        first_name = text.strip()
        if not first_name:
            QMessageBox.warning(self, "Error", "El nombre no puede estar vacío.")
            return
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            'UPDATE contacts SET first_name = ? WHERE id = ?', (first_name, contact_id)
        )
        conn.commit()
        conn.close()
        self.log_message("✓ Nombre de contacto actualizado")
        self.refresh_tables()

    def _edit_contact_last_name(self, contact_id):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT last_name FROM contacts WHERE id = ?', (contact_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return
        current = row[0] or ""
        text, ok = QInputDialog.getText(
            self, "Editar apellido(s)", "Apellido(s) del contacto:", text=current
        )
        if not ok:
            return
        last_name = text.strip()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            'UPDATE contacts SET last_name = ? WHERE id = ?', (last_name, contact_id)
        )
        conn.commit()
        conn.close()
        self.log_message("✓ Apellido(s) de contacto actualizado(s)")
        self.refresh_tables()
    
    def _edit_contact_email(self, contact_id):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT email FROM contacts WHERE id = ?', (contact_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return
        current = row[0]
        text, ok = QInputDialog.getText(
            self, "Editar correo", "Correo electrónico:", text=current
        )
        if not ok:
            return
        email = text.strip()
        if not email:
            QMessageBox.warning(self, "Error", "El correo no puede estar vacío.")
            return
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('UPDATE contacts SET email = ? WHERE id = ?', (email, contact_id))
            conn.commit()
            conn.close()
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Error", "Ese correo ya está registrado en otro contacto.")
            return
        self.log_message("✓ Correo de contacto actualizado")
        self.refresh_tables()
    
    def _edit_contact_phone(self, contact_id):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT phone FROM contacts WHERE id = ?', (contact_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return
        current = row[0] or ""
        dlg = QInputDialog(self)
        dlg.setWindowTitle("Editar teléfono")
        dlg.setLabelText("Teléfono (opcional, solo números):")
        dlg.setTextValue(current)
        le = dlg.findChild(QLineEdit)
        if le is not None:
            le.setValidator(_PHONE_DIGITS_VALIDATOR)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        phone = dlg.textValue().strip()
        if phone and not phone.isdigit():
            QMessageBox.warning(
                self, "Error", "El teléfono sólo puede contener números."
            )
            return
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('UPDATE contacts SET phone = ? WHERE id = ?', (phone, contact_id))
        conn.commit()
        conn.close()
        self.log_message("✓ Teléfono de contacto actualizado")
        self.refresh_tables()
    
    def _delete_contact(self, contact_id):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            'SELECT first_name, last_name FROM contacts WHERE id = ?', (contact_id,)
        )
        row = cursor.fetchone()
        if not row:
            conn.close()
            return
        display = contact_full_name(row[0], row[1]) or "contacto"
        conn.close()
        reply = QMessageBox.question(
            self,
            "Eliminar contacto",
            f"¿Eliminar el contacto «{display}»?\n"
            "Las citas vinculadas quedarán sin contacto asignado.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('UPDATE appointments SET contact_id = NULL WHERE contact_id = ?', (contact_id,))
        cursor.execute('DELETE FROM contacts WHERE id = ?', (contact_id,))
        conn.commit()
        conn.close()
        self.log_message(f"✓ Contacto eliminado: {display}")
        self.refresh_tables()
    
    def open_new_contact_dialog(self):
        """Abre el diálogo para crear un nuevo contacto."""
        dialog = NewContactDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.log_message("✓ Nuevo contacto agregado")
            self.refresh_tables()
    
    def open_template_editor(self):
        """Abre el editor de plantilla de correo."""
        dialog = TemplateEditorDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.log_message("✓ Plantilla de email actualizada")

    def open_contacts_dialog(self):
        """Abre la ventana emergente de contactos."""
        dialog = ContactsDialog(self)
        dialog.exec()

    def open_app_settings(self):
        """Abre preferencias generales de la aplicación."""
        dialog = SettingsDialog(self)
        dialog.exec()

    def open_help_dialog(self):
        """Abre la ventana de ayuda e información de PreCita."""
        dialog = HelpDialog(self)
        dialog.exec()

    def open_storage_dialog(self):
        """Abre la ventana con información del almacenamiento interno."""
        dialog = StorageDialog(self)
        dialog.exec()

    def close_google_session(self):
        """Cierra la sesión local de Google tras confirmación del usuario."""
        if not CREDENTIALS_PATH.exists():
            QMessageBox.information(
                self,
                "Sin sesión guardada",
                "No hay credenciales de Google guardadas en este equipo.",
            )
            return
        reply = QMessageBox.warning(
            self,
            "Cerrar sesión de Google",
            "Vas a cerrar la sesión de Google en este equipo.\n\n"
            "No podrás sincronizar el calendario ni enviar correos hasta iniciar sesión de nuevo.\n\n"
            "¿Seguro que deseas continuar?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        revoke_and_remove_google_credentials()
        QMessageBox.information(
            self,
            "Sesión cerrada",
            "La sesión local de Google se cerró correctamente.",
        )
        self.log_message("✓ Sesión de Google cerrada desde el menú principal")
        self.update_google_status_dot()
    
    def create_tray_icon(self):
        """Crea el icono en la bandeja del sistema."""
        self.tray_icon = QSystemTrayIcon(self)
        
        tray_menu = QMenu()
        show_action = tray_menu.addAction("Mostrar PreCita")
        show_action.triggered.connect(self.show)
        
        hide_action = tray_menu.addAction("Ocultar")
        hide_action.triggered.connect(self.hide)
        
        tray_menu.addSeparator()
        
        sync_action = tray_menu.addAction("Sincronizar ahora")
        sync_action.triggered.connect(self.sync_calendar)
        
        reminder_action = tray_menu.addAction("Enviar recordatorios")
        reminder_action.triggered.connect(self.send_reminders)
        
        tray_menu.addSeparator()
        
        exit_action = tray_menu.addAction("Salir")
        exit_action.triggered.connect(self.quit_app)
        
        icon_path = Path(__file__).parent / 'precita.ico'
        if icon_path.exists():
            icon = QIcon(str(icon_path))
            self.tray_icon.setIcon(icon)
            self.setWindowIcon(icon)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
    
    def quit_app(self):
        """Cierra la aplicación."""
        self.update_timer.stop()
        self.reminder_timer.stop()
        QApplication.quit()

    def activate_from_external_request(self):
        """Restaura y enfoca la ventana cuando se abre otra instancia."""
        self.setWindowState(self.windowState() & ~Qt.WindowState.WindowMinimized)
        self.show()
        self.raise_()
        self.activateWindow()
    
    def changeEvent(self, event):
        """Minimiza a la bandeja del sistema cuando se cierra."""
        if event.type() == 99:  # QEvent::WindowStateChange
            if self.windowState() & Qt.WindowState.WindowMinimized:
                self.hide()
                event.ignore()


def notify_existing_instance() -> bool:
    """Notifica a la instancia existente y devuelve True si ya está abierta."""
    socket_client = QLocalSocket()
    socket_client.connectToServer(SINGLE_INSTANCE_SERVER_NAME)
    if not socket_client.waitForConnected(250):
        return False
    socket_client.write(b"activate")
    socket_client.flush()
    socket_client.waitForBytesWritten(250)
    socket_client.disconnectFromServer()
    return True


def setup_single_instance_server(window: PreCitaMainWindow) -> QLocalServer | None:
    """Abre servidor local para recibir eventos de segunda ejecución."""
    server = QLocalServer(window)

    def _activate_window():
        while server.hasPendingConnections():
            client = server.nextPendingConnection()
            if client is not None:
                client.disconnectFromServer()
            window.activate_from_external_request()

    server.newConnection.connect(_activate_window)

    if server.listen(SINGLE_INSTANCE_SERVER_NAME):
        return server

    # Posible socket huérfano tras cierre abrupto.
    QLocalServer.removeServer(SINGLE_INSTANCE_SERVER_NAME)
    if server.listen(SINGLE_INSTANCE_SERVER_NAME):
        return server
    return None

# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================

if __name__ == '__main__':
    app = QApplication(sys.argv)
    if notify_existing_instance():
        sys.exit(0)
    if not prepare_database_for_runtime():
        sys.exit(0)
    init_database()
    app.aboutToQuit.connect(finalize_database_encryption_on_exit)
    if sys.platform == "win32":
        # Autorrepara la ruta de inicio si el EXE/script se movió.
        sync_windows_startup_with_settings()
    _th = get_setting("theme", DEFAULT_THEME) or DEFAULT_THEME
    _display_scale = get_display_scale_percent()
    apply_app_appearance(app, _th, _display_scale)
    window = PreCitaMainWindow()
    window.single_instance_server = setup_single_instance_server(window)
    window.showMaximized()
    sys.exit(app.exec())
