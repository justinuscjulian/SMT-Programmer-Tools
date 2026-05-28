from PySide6.QtCore import QEasingCurve, QPropertyAnimation, Qt, QThreadPool
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QFrame, QGraphicsOpacityEffect, QHBoxLayout, QMainWindow, QStackedWidget, QVBoxLayout, QWidget

from ui.pages.bom_compare_page import BomComparePage
from ui.pages.help_page import HelpPage
from ui.pages.history_page import HistoryPage
from ui.pages.machine_compare_page import MachineComparePage
from ui.pages.other_tools_page import OtherToolsPage
from ui.pages.plan_page import PlanPage
from ui.sidebar import Sidebar
from ui.top_bar import TopBar
from utils.paths import resource_path


class MainWindow(QMainWindow):
    def __init__(self, theme_manager, parent=None):
        super().__init__(parent)
        self.theme_manager = theme_manager
        self.thread_pool = QThreadPool.globalInstance()
        self.thread_pool.setMaxThreadCount(max(2, min(8, self.thread_pool.maxThreadCount())))
        self.titles = {
            "compare": "BOM Comparison",
            "machine": "Machine Data Audit",
            "plan": "PLAN",
            "history": "History & Logs",
            "other": "Other Tools",
            "help": "Usage Guide",
        }
        self.pages = {}
        self._animations = []
        self._build_window()

    def _build_window(self):
        self.setWindowTitle("SMT Programmer Tools")
        self.resize(1500, 900)
        self.setMinimumSize(1200, 780)

        icon_path = resource_path("assets/app_logo.ico")
        icon = QIcon(icon_path)
        if not icon.isNull():
            self.setWindowIcon(icon)

        app_shell = QWidget()
        app_shell.setObjectName("AppRoot")
        root = QHBoxLayout(app_shell)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(16)

        self.sidebar = Sidebar(self.theme_manager)
        self.sidebar.nav_selected.connect(self.show_page)
        root.addWidget(self.sidebar)

        main_area = QFrame()
        main_area.setObjectName("MainArea")
        main_layout = QVBoxLayout(main_area)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(16)

        self.top_bar = TopBar(self.theme_manager)
        self.top_bar.help_requested.connect(lambda: self.show_page("help"))
        main_layout.addWidget(self.top_bar)

        content = QWidget()
        content.setObjectName("ContentArea")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        self.stack = QStackedWidget()
        self.pages = {
            "compare": BomComparePage(self.thread_pool, self.theme_manager),
            "machine": MachineComparePage(self.thread_pool, self.theme_manager),
            "plan": PlanPage(self.thread_pool, self.theme_manager),
            "history": HistoryPage(self.thread_pool, self.theme_manager),
            "other": OtherToolsPage(self.thread_pool, self.theme_manager),
            "help": HelpPage(self.theme_manager),
        }
        for page in self.pages.values():
            self.stack.addWidget(page)

        content_layout.addWidget(self.stack, 1)
        main_layout.addWidget(content, 1)
        root.addWidget(main_area, 1)
        self.app_shell = app_shell
        self.setCentralWidget(self.app_shell)

        self.show_page("compare", animate=False)

    def show_page(self, view_id, animate=True):
        page = self.pages.get(view_id)
        if page is None:
            return
        self.stack.setCurrentWidget(page)
        if animate:
            self._fade_widget(page, duration=180)
        self.top_bar.set_title(self.titles.get(view_id, "Dashboard"))
        self.sidebar.active_id = view_id
        self.sidebar._update_button_styles()
        if view_id == "history":
            page.refresh_history()

    def _fade_widget(self, widget, duration=180):
        effect = QGraphicsOpacityEffect(widget)
        effect.setOpacity(0.0)
        widget.setGraphicsEffect(effect)

        animation = QPropertyAnimation(effect, b"opacity", self)
        animation.setDuration(duration)
        animation.setStartValue(0.0)
        animation.setEndValue(1.0)
        animation.setEasingCurve(QEasingCurve.OutCubic)
        animation.finished.connect(lambda w=widget, a=animation: self._finish_animation(w, a))
        self._animations.append(animation)
        animation.start()

    def _finish_animation(self, widget, animation):
        widget.setGraphicsEffect(None)
        if animation in self._animations:
            self._animations.remove(animation)
