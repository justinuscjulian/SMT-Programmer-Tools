from PySide6.QtCore import QObject, QSettings, Signal

from .tokens import THEMES


def build_qss(theme):
    return f"""
    QMainWindow, QWidget#AppRoot {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["window_2"]},
            stop:0.44 {theme["window"]},
            stop:1 {theme["window_2"]}
        );
    }}
    QWidget {{
        background: transparent;
        color: {theme["text"]};
        font-family: "Segoe UI";
        font-size: 10pt;
    }}
    QFrame#Card, QWidget#Card {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:0.56 {theme["glass"]},
            stop:1 {theme["surface"]}
        );
        border: 1px solid {theme["border_soft"]};
        border-radius: 22px;
    }}
    QFrame#SubPanel {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:0.58 {theme["glass"]},
            stop:1 {theme["surface"]}
        );
        border: 1px solid {theme["border_soft"]};
        border-radius: 18px;
    }}
    QFrame#TopBar, QFrame#Sidebar {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["glass"]}
        );
        border: 1px solid {theme["border_soft"]};
        border-radius: 24px;
    }}
    QFrame#MainArea {{
        background: transparent;
        border: none;
    }}
    QWidget#ContentArea {{
        background: transparent;
        border: none;
    }}
    QLabel#TitleLabel {{
        font-size: 21pt;
        font-weight: 700;
    }}
    QLabel#SectionTitle {{
        font-size: 12pt;
        font-weight: 700;
    }}
    QLabel#MutedLabel {{
        color: {theme["muted"]};
    }}
    QLabel#GuideMiniTitle {{
        color: {theme["accent"]};
        font-size: 10pt;
        font-weight: 800;
        padding-top: 4px;
    }}
    QLabel#GuideBody {{
        color: {theme["text"]};
        line-height: 1.35;
    }}
    QPushButton, QToolButton {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:0.52 {theme["surface_high"]},
            stop:1 {theme["glass"]}
        );
        border: 1px solid {theme["border_soft"]};
        border-radius: 15px;
        padding: 9px 14px;
        font-weight: 600;
    }}
    QPushButton:hover, QToolButton:hover {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["surface_hover"]},
            stop:0.52 {theme["glass_strong"]},
            stop:1 {theme["accent_soft"]}
        );
        border-color: {theme["border"]};
    }}
    QPushButton:pressed, QToolButton:pressed {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["accent_soft"]},
            stop:1 {theme["glass"]}
        );
    }}
    QPushButton:disabled, QToolButton:disabled {{
        color: {theme["muted"]};
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass"]},
            stop:1 {theme["surface"]}
        );
    }}
    QPushButton#PrimaryButton {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["accent"]},
            stop:1 {theme["accent_hover"]}
        );
        color: #ffffff;
        border-color: {theme["border_soft"]};
    }}
    QPushButton#PrimaryButton:hover {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["accent_hover"]},
            stop:1 {theme["accent"]}
        );
    }}
    QPushButton#DangerButton {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["glass"]}
        );
        color: {theme["red"]};
        border-color: {theme["border"]};
    }}
    QPushButton#SuccessButton {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["green"]},
            stop:1 {theme["accent"]}
        );
        color: #ffffff;
        border-color: {theme["border_soft"]};
    }}
    QPushButton#SegmentedButton {{
        border-radius: 16px;
        padding: 8px 18px;
    }}
    QPushButton#SegmentedButton:checked {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["accent"]},
            stop:1 {theme["accent_hover"]}
        );
        color: #ffffff;
        border-color: {theme["accent"]};
    }}
    QPushButton#ThemeToggle {{
        border-radius: 18px;
        padding: 8px 16px;
        min-width: 126px;
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:0.55 {theme["surface_high"]},
            stop:1 {theme["glass"]}
        );
        border: 1px solid {theme["border_soft"]};
    }}
    QPushButton#HelpButton {{
        border-radius: 16px;
        padding: 7px;
        min-width: 34px;
        max-width: 34px;
        min-height: 34px;
        max-height: 34px;
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["accent_soft"]}
        );
        border: 1px solid {theme["border_soft"]};
    }}
    QPushButton#HelpButton:hover {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["surface_hover"]},
            stop:1 {theme["accent_soft"]}
        );
        border-color: {theme["accent"]};
    }}
    QLineEdit, QComboBox, QTextEdit, QPlainTextEdit {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:0.62 {theme["surface_high"]},
            stop:1 {theme["glass"]}
        );
        border: 1px solid {theme["border_soft"]};
        border-radius: 15px;
        padding: 8px 11px;
        selection-background-color: {theme["accent"]};
    }}
    QComboBox::drop-down {{
        border: none;
        width: 24px;
    }}
    QComboBox QAbstractItemView {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["glass"]}
        );
        border: 1px solid {theme["border_soft"]};
        border-radius: 12px;
        padding: 6px;
        selection-background-color: {theme["accent_soft"]};
        selection-color: {theme["accent"]};
    }}
    QTableView {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:0.58 {theme["glass"]},
            stop:1 {theme["surface"]}
        );
        alternate-background-color: {theme["surface"]};
        border: 1px solid {theme["border_soft"]};
        border-radius: 16px;
        gridline-color: {theme["border"]};
        selection-background-color: {theme["accent_soft"]};
        selection-color: {theme["text"]};
    }}
    QHeaderView::section {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:0,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["surface"]}
        );
        color: {theme["muted"]};
        border: none;
        border-bottom: 1px solid {theme["border"]};
        padding: 10px;
        font-weight: 700;
    }}
    QScrollBar:vertical, QScrollBar:horizontal {{
        background: transparent;
        border: none;
        margin: 2px;
    }}
    QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["border"]}
        );
        border-radius: 5px;
        min-height: 28px;
        min-width: 28px;
    }}
    QSplitter::handle {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 transparent,
            stop:0.5 {theme["glass"]},
            stop:1 transparent
        );
        border: none;
    }}
    QProgressBar {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["glass"]}
        );
        border: 1px solid {theme["border_soft"]};
        border-radius: 10px;
        height: 8px;
        text-align: center;
    }}
    QProgressBar::chunk {{
        background: {theme["accent"]};
        border-radius: 10px;
    }}
    QToolButton#NavButton {{
        text-align: left;
        border: 1px solid transparent;
        border-radius: 18px;
        padding: 10px 12px;
        min-height: 42px;
        background: transparent;
    }}
    QToolButton#NavButton[collapsed="false"] {{
        padding-left: 22px;
        padding-right: 16px;
    }}
    QToolButton#NavButton[collapsed="true"] {{
        padding-left: 12px;
        padding-right: 12px;
        text-align: center;
    }}
    QToolButton#NavButton:hover {{
        background: transparent;
        border-color: transparent;
    }}
    QToolButton#NavButton[active="true"] {{
        background: transparent;
        color: {theme["accent"]};
        border-color: transparent;
    }}
    QToolButton#SidebarToggle {{
        border-radius: 18px;
        min-width: 38px;
        min-height: 38px;
        padding: 6px;
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["glass"]}
        );
    }}
    QToolButton#ToolMenuButton {{
        text-align: left;
        border-radius: 18px;
        padding: 14px 16px;
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:0.55 {theme["surface_high"]},
            stop:1 {theme["glass"]}
        );
        border: 1px solid {theme["border_soft"]};
        font-weight: 700;
    }}
    QToolButton#ToolMenuButton:hover {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["surface_hover"]},
            stop:1 {theme["accent_soft"]}
        );
        border-color: {theme["accent"]};
    }}
    QLabel#BrandChip {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["accent_soft"]}
        );
        color: {theme["accent"]};
        border: 1px solid {theme["border_soft"]};
        border-radius: 16px;
        padding: 8px 12px;
        font-weight: 800;
    }}
    QLabel#StatusBadge {{
        border-radius: 13px;
        padding: 6px 12px;
        font-weight: 700;
    }}
    QLabel#StatusBadge[status="ADD"] {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["add_bg"]}
        );
        color: {theme["add_fg"]};
    }}
    QLabel#StatusBadge[status="CNG"] {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["cng_bg"]}
        );
        color: {theme["cng_fg"]};
    }}
    QLabel#StatusBadge[status="MOVE"] {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["cng_bg"]}
        );
        color: {theme["cng_fg"]};
    }}
    QLabel#StatusBadge[status="DEL"] {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["del_bg"]}
        );
        color: {theme["del_fg"]};
    }}
    QLabel#StatusBadge[status="INFO"] {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["accent_soft"]}
        );
        color: {theme["accent"]};
    }}
    QMenu {{
        background: {theme["glass_strong"]};
        color: {theme["text"]};
        border: 1px solid {theme["border_soft"]};
        border-radius: 14px;
        padding: 8px;
    }}
    QMenu::item {{
        padding: 8px 24px;
        border-radius: 10px;
    }}
    QMenu::item:selected {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["accent_soft"]}
        );
        color: {theme["accent"]};
    }}
    QMessageBox, QDialog {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:0.58 {theme["glass"]},
            stop:1 {theme["surface"]}
        );
        color: {theme["text"]};
    }}
    QMessageBox QLabel, QDialog QLabel {{
        background: transparent;
        color: {theme["text"]};
    }}
    QMessageBox QPushButton, QDialog QPushButton {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["glass_strong"]},
            stop:1 {theme["surface_high"]}
        );
        color: {theme["text"]};
        border: 1px solid {theme["border"]};
        border-radius: 12px;
        padding: 8px 18px;
        min-width: 82px;
    }}
    QMessageBox QPushButton:hover, QDialog QPushButton:hover {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0 {theme["surface_hover"]},
            stop:1 {theme["accent_soft"]}
        );
        border-color: {theme["accent"]};
    }}
    """


class ThemeManager(QObject):
    changed = Signal(dict)

    def __init__(self, app):
        super().__init__()
        self.app = app
        self.settings = QSettings("SMTTools", "BomComparatorQt")
        self.mode = self.settings.value("theme", "light")
        if self.mode not in THEMES:
            self.mode = "light"

    @property
    def theme(self):
        return THEMES[self.mode]

    def apply(self):
        self.app.setStyleSheet(build_qss(self.theme))
        self.settings.setValue("theme", self.mode)
        self.changed.emit(self.theme)

    def toggle(self):
        self.mode = "dark" if self.mode == "light" else "light"
        self.apply()
