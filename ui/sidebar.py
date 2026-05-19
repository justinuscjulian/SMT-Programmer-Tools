from PySide6.QtCore import Property, QEasingCurve, QPropertyAnimation, QSize, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QFrame, QGraphicsDropShadowEffect, QHBoxLayout, QLabel, QSizePolicy, QToolButton, QVBoxLayout

from ui.icons import make_icon
from ui.liquid_nav_button import LiquidNavButton


class Sidebar(QFrame):
    nav_selected = Signal(str)

    def __init__(self, theme_manager, parent=None):
        super().__init__(parent)
        self.setObjectName("Sidebar")
        self.theme_manager = theme_manager
        self.expanded_width = 240
        self.collapsed_width = 82
        self._sidebar_width = self.expanded_width
        self.is_collapsed = False
        self.active_id = "compare"
        self.buttons = {}
        self._width_animation = QPropertyAnimation(self, b"sidebarWidth", self)
        self._width_animation.setDuration(320)
        self._width_animation.setEasingCurve(QEasingCurve.OutCubic)

        self.setFixedWidth(self.expanded_width)
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(36)
        shadow.setOffset(0, 12)
        shadow.setColor(QColor(15, 23, 42, 38))
        self.setGraphicsEffect(shadow)
        self._build_ui()
        self.theme_manager.changed.connect(lambda _: self._refresh_icons())

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(14)

        top = QHBoxLayout()
        self.toggle_btn = QToolButton()
        self.toggle_btn.setObjectName("SidebarToggle")
        self.toggle_btn.setIconSize(QSize(22, 22))
        self.toggle_btn.setToolTip("Toggle sidebar")
        self.toggle_btn.clicked.connect(self.toggle)
        self.logo_label = QLabel("SMT Tools")
        self.logo_label.setObjectName("BrandChip")
        top.addWidget(self.toggle_btn)
        top.addWidget(self.logo_label)
        top.addStretch(1)
        root.addLayout(top)

        self._add_nav(root, "compare", "compare", "BOM Comparison")
        self._add_nav(root, "machine", "machine", "Machine Data Audit")
        self._add_nav(root, "other", "tools", "Other Tools")
        root.addStretch(1)
        self._add_nav(root, "history", "history", "History & Logs")

        credit = QLabel("Created by:\nJustinus CJ")
        credit.setObjectName("MutedLabel")
        credit.setAlignment(Qt.AlignLeft)
        root.addWidget(credit)
        self.credit = credit
        self._refresh_icons()
        self._update_button_styles()

    def _add_nav(self, layout, view_id, icon_name, label):
        button = LiquidNavButton(self.theme_manager)
        button.setObjectName("NavButton")
        button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        button.setIconSize(QSize(22, 22))
        button.setText(label)
        button.setToolTip(label)
        button.setMinimumHeight(48)
        button.setProperty("collapsed", self.is_collapsed)
        button.clicked.connect(lambda checked=False, vid=view_id: self.select(vid))
        self.buttons[view_id] = {"button": button, "icon": icon_name, "label": label}
        layout.addWidget(button)

    def select(self, view_id):
        self.active_id = view_id
        self._update_button_styles()
        self.nav_selected.emit(view_id)

    def toggle(self):
        start = self.width()
        target = self.collapsed_width if not self.is_collapsed else self.expanded_width
        self.is_collapsed = not self.is_collapsed
        self.logo_label.setVisible(not self.is_collapsed)
        self.credit.setVisible(not self.is_collapsed)
        self._update_button_text()

        self._width_animation.stop()
        self._width_animation.setStartValue(start)
        self._width_animation.setEndValue(target)
        self._width_animation.start()

    def _update_button_text(self):
        for data in self.buttons.values():
            button = data["button"]
            button.setProperty("collapsed", self.is_collapsed)
            if self.is_collapsed:
                button.setText("")
                button.setToolButtonStyle(Qt.ToolButtonIconOnly)
            else:
                button.setText(data["label"])
                button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
            self._polish_button(button)

    def _update_button_styles(self):
        for view_id, data in self.buttons.items():
            active = view_id == self.active_id
            button = data["button"]
            button.setProperty("active", active)
            button.setProperty("collapsed", self.is_collapsed)
            button.set_active(active)
            self._polish_button(button)
        self._refresh_icons()

    def _polish_button(self, button):
        button.style().unpolish(button)
        button.style().polish(button)
        button.update()

    def _refresh_icons(self):
        theme = self.theme_manager.theme
        self.toggle_btn.setIcon(make_icon("menu", theme["text"]))
        for view_id, data in self.buttons.items():
            color = theme["accent"] if view_id == self.active_id else theme["muted"]
            data["button"].setIcon(make_icon(data["icon"], color))

    def get_sidebar_width(self):
        return self._sidebar_width

    def set_sidebar_width(self, width):
        self._sidebar_width = width
        self.setFixedWidth(width)

    sidebarWidth = Property(int, get_sidebar_width, set_sidebar_width)
