from datetime import datetime

from PySide6.QtCore import QSize, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QFrame, QGraphicsDropShadowEffect, QHBoxLayout, QLabel, QPushButton

from ui.icons import make_icon


class TopBar(QFrame):
    def __init__(self, theme_manager, parent=None):
        super().__init__(parent)
        self.setObjectName("TopBar")
        self.theme_manager = theme_manager
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(30)
        shadow.setOffset(0, 10)
        shadow.setColor(QColor(15, 23, 42, 28))
        self.setGraphicsEffect(shadow)
        self._build_ui()
        self.theme_manager.changed.connect(lambda _: self.update_theme_button())

    def _build_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(22, 14, 22, 14)
        layout.setSpacing(14)

        self.title_label = QLabel("BOM Comparison")
        self.title_label.setObjectName("TitleLabel")
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedLabel")
        self.clock_label = QLabel("")
        self.clock_label.setObjectName("MutedLabel")
        self.theme_btn = QPushButton("")
        self.theme_btn.setObjectName("ThemeToggle")
        self.theme_btn.setIconSize(QSize(20, 20))
        self.theme_btn.clicked.connect(self.theme_manager.toggle)

        layout.addWidget(self.title_label)
        layout.addWidget(self.status_label)
        layout.addStretch(1)
        layout.addWidget(self.clock_label)
        layout.addWidget(self.theme_btn)

        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_clock)
        self.clock_timer.start(1000)
        self.update_clock()
        self.update_theme_button()

    def set_title(self, title):
        self.title_label.setText(title)

    def set_status(self, status):
        self.status_label.setText(status)

    def update_clock(self):
        self.clock_label.setText(datetime.now().strftime("%d/%m/%Y  %H:%M:%S"))

    def update_theme_button(self):
        theme = self.theme_manager.theme
        if self.theme_manager.mode == "dark":
            self.theme_btn.setText("Light")
            self.theme_btn.setIcon(make_icon("sun", theme["orange"]))
        else:
            self.theme_btn.setText("Dark")
            self.theme_btn.setIcon(make_icon("moon", theme["accent"]))
