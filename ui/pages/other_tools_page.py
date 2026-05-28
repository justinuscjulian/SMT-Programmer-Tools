from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import QHBoxLayout, QLabel, QPushButton, QStackedWidget, QToolButton, QVBoxLayout, QWidget

from ui.icons import make_icon
from ui.pages.all_in_one_page import AllInOneComparatorPage
from ui.pages.insert_point_page import InsertPointPage
from ui.pages.new_pcb_excel_page import NewPcbExcelPage
from ui.pages.worksheet_bom_compare_page import WorksheetBomComparePage
from ui.pages.worksheet_page import WorksheetComparatorPage
from widgets.card import Card


class OtherToolsPage(QWidget):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(parent)
        self.thread_pool = thread_pool
        self.theme_manager = theme_manager
        self.tool_buttons = []
        self._build_ui()
        self.theme_manager.changed.connect(lambda _: self._refresh_icons())

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        self.stack = QStackedWidget()
        self.stack.addWidget(self._build_menu_page())
        self.stack.addWidget(self._build_worksheet_page())
        self.stack.addWidget(self._build_worksheet_bom_page())
        self.stack.addWidget(self._build_all_in_one_page())
        self.stack.addWidget(self._build_new_pcb_page())
        self.stack.addWidget(self._build_insert_point_page())
        root.addWidget(self.stack, 1)
        self._refresh_icons()

    def _build_menu_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        header = QHBoxLayout()
        subtitle = QLabel("Pilih tool yang ingin digunakan")
        subtitle.setObjectName("MutedLabel")
        header.addWidget(subtitle)
        header.addStretch(1)
        layout.addLayout(header)

        menu_card = Card()
        card_title = QLabel("List Menu")
        card_title.setObjectName("SectionTitle")
        menu_card.layout.addWidget(card_title)

        self.worksheet_button = self._create_tool_button(
            "Worksheet Comparator",
            "Worksheet vs CRB Verification",
            "worksheet",
            self.open_worksheet,
        )
        menu_card.layout.addWidget(self.worksheet_button)
        self.worksheet_bom_button = self._create_tool_button(
            "Worksheet vs BOM Comparator",
            "Compare Worksheet dengan BOM .tsv",
            "compare",
            self.open_worksheet_bom,
        )
        menu_card.layout.addWidget(self.worksheet_bom_button)
        self.all_in_one_button = self._create_tool_button(
            "All In One Comparator",
            "Compare NPM, BM, dan BOM dari satu layar",
            "all_in_one",
            self.open_all_in_one,
        )
        menu_card.layout.addWidget(self.all_in_one_button)
        self.new_pcb_button = self._create_tool_button(
            "NEW PCB Excel Creator",
            "Generate Excel program SMT",
            "worksheet",
            self.open_new_pcb,
        )
        menu_card.layout.addWidget(self.new_pcb_button)
        self.insert_point_button = self._create_tool_button(
            "Get Insert Point",
            "Ambil data Insert Point dari folder PCB",
            "insert_point",
            self.open_insert_point,
        )
        menu_card.layout.addWidget(self.insert_point_button)
        menu_card.layout.addStretch(1)
        layout.addWidget(menu_card, 1)

        return page

    def _build_worksheet_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.back_button = QPushButton("Back to Tools")
        self.back_button.clicked.connect(self.open_menu)
        title = QLabel("Worksheet Comparator")
        title.setObjectName("TitleLabel")
        top.addWidget(self.back_button)
        top.addWidget(title)
        top.addStretch(1)
        layout.addLayout(top)

        self.worksheet_page = WorksheetComparatorPage(self.thread_pool, self.theme_manager)
        layout.addWidget(self.worksheet_page, 1)
        return page

    def _build_worksheet_bom_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.back_worksheet_bom_button = QPushButton("Back to Tools")
        self.back_worksheet_bom_button.clicked.connect(self.open_menu)
        title = QLabel("Worksheet vs BOM Comparator")
        title.setObjectName("TitleLabel")
        top.addWidget(self.back_worksheet_bom_button)
        top.addWidget(title)
        top.addStretch(1)
        layout.addLayout(top)

        self.worksheet_bom_page = WorksheetBomComparePage(self.thread_pool, self.theme_manager)
        layout.addWidget(self.worksheet_bom_page, 1)
        return page

    def _build_all_in_one_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.back_all_in_one_button = QPushButton("Back to Tools")
        self.back_all_in_one_button.clicked.connect(self.open_menu)
        title = QLabel("All In One Comparator")
        title.setObjectName("TitleLabel")
        top.addWidget(self.back_all_in_one_button)
        top.addWidget(title)
        top.addStretch(1)
        layout.addLayout(top)

        self.all_in_one_page = AllInOneComparatorPage(self.thread_pool, self.theme_manager)
        layout.addWidget(self.all_in_one_page, 1)
        return page

    def _build_new_pcb_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.back_new_pcb_button = QPushButton("Back to Tools")
        self.back_new_pcb_button.clicked.connect(self.open_menu)
        title = QLabel("NEW PCB Excel Creator")
        title.setObjectName("TitleLabel")
        top.addWidget(self.back_new_pcb_button)
        top.addWidget(title)
        top.addStretch(1)
        layout.addLayout(top)

        self.new_pcb_page = NewPcbExcelPage(self.thread_pool, self.theme_manager)
        layout.addWidget(self.new_pcb_page, 1)
        return page

    def _build_insert_point_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.back_insert_point_button = QPushButton("Back to Tools")
        self.back_insert_point_button.clicked.connect(self.open_menu)
        title = QLabel("Get Insert Point")
        title.setObjectName("TitleLabel")
        top.addWidget(self.back_insert_point_button)
        top.addWidget(title)
        top.addStretch(1)
        layout.addLayout(top)

        self.insert_point_page = InsertPointPage(self.thread_pool, self.theme_manager)
        layout.addWidget(self.insert_point_page, 1)
        return page

    def _create_tool_button(self, title, subtitle, icon_name, callback):
        button = QToolButton()
        button.setObjectName("ToolMenuButton")
        button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        button.setIconSize(QSize(30, 30))
        button.setMinimumHeight(72)
        button.setText(f"{title}\n{subtitle}")
        button.setToolTip(title)
        button.clicked.connect(callback)
        button._tool_icon_name = icon_name
        self.tool_buttons.append(button)
        return button

    def open_menu(self):
        self.stack.setCurrentIndex(0)

    def open_worksheet(self):
        self.stack.setCurrentIndex(1)

    def open_all_in_one(self):
        self.stack.setCurrentIndex(3)

    def open_new_pcb(self):
        self.stack.setCurrentIndex(4)

    def open_insert_point(self):
        self.stack.setCurrentIndex(5)

    def open_worksheet_bom(self):
        self.stack.setCurrentIndex(2)

    def _refresh_icons(self):
        theme = self.theme_manager.theme
        for button in self.tool_buttons:
            button.setIcon(make_icon(button._tool_icon_name, theme["accent"]))
        if hasattr(self, "back_button"):
            self.back_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_worksheet_bom_button"):
            self.back_worksheet_bom_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_all_in_one_button"):
            self.back_all_in_one_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_new_pcb_button"):
            self.back_new_pcb_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_insert_point_button"):
            self.back_insert_point_button.setIcon(make_icon("arrow_left", theme["text"]))
