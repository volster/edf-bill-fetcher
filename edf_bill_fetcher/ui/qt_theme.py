"""Desktop design tokens and an accessible, consistent Qt widget theme."""

from PySide6.QtGui import QFont, QFontDatabase, QPalette
from PySide6.QtWidgets import QApplication

INK = "#18283F"
MUTED = "#637086"
ACCENT = "#AF451E"
BACKGROUND = "#F4F6FA"

STYLESHEET = """
QWidget { color: #18283F; font-size: 13px; }
QMainWindow, QScrollArea, QStackedWidget, QWidget#page { background: #F4F6FA; }
QScrollArea { border: none; }
QWidget#sidebar { background: #14263F; }
QWidget#sidebar QLabel { color: #B4C2D5; background: transparent; }
QWidget#sidebar QLabel#brand { color: #FFFFFF; font-size: 21px; font-weight: 700; }
QWidget#sidebar QLabel#eyebrow { color: #F7B38E; font-size: 11px; font-weight: 700; }
QPushButton#nav { text-align: left; border: none; color: #CCD6E3; background: transparent;
    padding: 13px 14px; border-radius: 7px; }
QPushButton#nav:hover { background: #223B58; }
QPushButton#nav:checked { color: #FFFFFF; background: #2C4562; font-weight: 600; }
QPushButton#nav:focus { border: 2px solid #F7B38E; padding: 11px 12px; }
QLabel#eyebrow { color: #AF451E; font-size: 11px; font-weight: 700; }
QLabel#title { font-size: 28px; font-weight: 700; color: #18283F; }
QLabel#subtitle, QLabel#hint { color: #637086; background: transparent; }
QLabel#sectionTitle { font-size: 16px; font-weight: 600; }
QLabel#fieldLabel { font-weight: 600; }
QFrame#card { background: white; border: 1px solid #DCE3EC; border-radius: 10px; }
QFrame#card QLabel, QFrame#card QWidget#row { background: transparent; }
QFrame#footer { background: white; border-top: 1px solid #DCE3EC; }
QLabel#status { font-weight: 600; font-size: 14px; }
QLabel#feedback { background: #EEF3FA; border: 1px solid #D6E1EE; padding: 12px;
    border-radius: 7px; color: #284766; }
QLabel#feedback[severity="success"] { background: #EDF7F1; border-color: #BDD9C9; color: #205D40; }
QLabel#feedback[severity="warning"] { background: #FFF7E8; border-color: #E8D2A4; color: #78501C; }
QLabel#feedback[severity="error"] { background: #FFF0EF; border-color: #EAC1BD; color: #8B302A; }
QLineEdit, QDoubleSpinBox { background: #FFFFFF; border: 1px solid #C6CFDC;
    border-radius: 6px; padding: 8px 10px; min-height: 20px; selection-background-color: #D6E5F5; }
QLineEdit:focus, QDoubleSpinBox:focus { border: 2px solid #3267A6; padding: 7px 9px; }
QLineEdit:disabled, QDoubleSpinBox:disabled { background: #F1F3F6; color: #8B95A5; }
QPushButton { background: white; border: 1px solid #C6CFDC; border-radius: 6px;
    padding: 9px 14px; font-weight: 600; min-height: 20px; }
QPushButton:hover { background: #F0F4FA; border-color: #91A4BC; }
QPushButton:pressed { background: #E4EBF4; }
QPushButton:focus { border: 2px solid #3267A6; padding: 8px 13px; }
QPushButton:disabled { background: #F1F3F6; color: #8792A4; border-color: #E1E6EE; }
QPushButton#primary { background: #AF451E; border-color: #AF451E; color: white; }
QPushButton#primary:hover { background: #963916; border-color: #963916; }
QPushButton#primary:focus { border: 2px solid #18283F; }
QPushButton#primary:disabled { background: #E7EBF1; border-color: #E7EBF1; color: #68758A; }
QPushButton#textButton { color: #3267A6; border: none; background: transparent; padding: 4px 0; }
QPushButton#textButton:hover { color: #18283F; text-decoration: underline; }
QPushButton#textButton:focus { border: 2px solid #3267A6; padding: 2px 0; }
QCheckBox, QRadioButton { spacing: 9px; padding: 5px 0; background: transparent; }
QCheckBox:disabled, QRadioButton:disabled { color: #8792A4; }
QCheckBox::indicator { width: 16px; height: 16px; }
QCheckBox::indicator:unchecked { border: 1px solid #8493A7; border-radius: 3px; background: white; }
QCheckBox::indicator:unchecked:disabled { border-color: #C6CFDC; background: #F1F3F6; }
QRadioButton::indicator { width: 17px; height: 17px; }
QRadioButton::indicator:unchecked { border: 1px solid #8493A7; border-radius: 9px; background: white; }
QProgressBar { border: none; background: #E8EDF4; border-radius: 2px; max-height: 5px; }
QProgressBar::chunk { background: #AF451E; border-radius: 2px; }
QScrollBar:vertical { background: transparent; width: 10px; margin: 3px; }
QScrollBar::handle:vertical { background: #BFCADA; border-radius: 3px; min-height: 30px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QDialog { background: #F4F6FA; }
QToolTip { background: #18283F; color: white; padding: 7px; border: none; }
"""


def apply_theme(application: QApplication) -> None:
    """Use scalable system fonts and a neutral palette on every platform."""
    application.setStyle("Fusion")
    families = QFontDatabase.families()
    family = next(
        (name for name in ("Segoe UI", "Inter", "Noto Sans", "DejaVu Sans") if name in families),
        application.font().family(),
    )
    application.setFont(QFont(family, 10))
    palette = application.palette()
    palette.setColor(QPalette.ColorRole.Window, BACKGROUND)
    palette.setColor(QPalette.ColorRole.WindowText, INK)
    palette.setColor(QPalette.ColorRole.Base, "#FFFFFF")
    palette.setColor(QPalette.ColorRole.Text, INK)
    palette.setColor(QPalette.ColorRole.Highlight, "#3267A6")
    palette.setColor(QPalette.ColorRole.HighlightedText, "#FFFFFF")
    application.setPalette(palette)
    application.setStyleSheet(STYLESHEET)
