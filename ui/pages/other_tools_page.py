from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import QGridLayout, QHBoxLayout, QLabel, QPushButton, QStackedWidget, QToolButton, QVBoxLayout, QWidget

from ui.icons import make_icon
from ui.pages.all_in_one_page import AllInOneComparatorPage
from ui.pages.component_usage_finder_page import ComponentUsageFinderPage
from ui.pages.common_feeder_reuse_page import CommonFeederReusePage
from ui.pages.feeder_mapping_page import FeederMappingPage
from ui.pages.insert_point_page import InsertPointPage
from ui.pages.model_feeder_group_page import ModelFeederGroupPage
from ui.pages.new_pcb_excel_page import NewPcbExcelPage
from ui.pages.used_part_component_page import UsedPartComponentPage
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
        self.stack.addWidget(self._build_feeder_mapping_page())
        self.stack.addWidget(self._build_all_in_one_page())
        self.stack.addWidget(self._build_new_pcb_page())
        self.stack.addWidget(self._build_insert_point_page())
        self.stack.addWidget(self._build_used_part_component_page())
        self.stack.addWidget(self._build_component_usage_page())
        self.stack.addWidget(self._build_common_feeder_reuse_page())
        self.stack.addWidget(self._build_model_feeder_group_page())
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

        tools_layout = QGridLayout()
        tools_layout.setHorizontalSpacing(12)
        tools_layout.setVerticalSpacing(12)
        for column in range(4):
            tools_layout.setColumnStretch(column, 1)

        tools = [
            (
                "worksheet_button",
                "Worksheet Comparator",
                "Worksheet vs CRB Verification",
                "worksheet",
                self.open_worksheet,
            ),
            (
                "worksheet_bom_button",
                "Worksheet vs BOM Comparator",
                "Compare Worksheet dengan BOM .tsv",
                "compare",
                self.open_worksheet_bom,
            ),
            (
                "feeder_mapping_button",
                "Feeder Mapping Generator",
                "Convert export TXT mesin NPM ke Excel",
                "feeder_mapping",
                self.open_feeder_mapping,
            ),
            (
                "all_in_one_button",
                "All In One Comparator",
                "Compare NPM, BM, dan BOM dari satu layar",
                "all_in_one",
                self.open_all_in_one,
            ),
            (
                "new_pcb_button",
                "NEW PCB Excel Creator",
                "Generate Excel program SMT",
                "worksheet",
                self.open_new_pcb,
            ),
            (
                "insert_point_button",
                "Get Insert Point",
                "Ambil data Insert Point dari folder PCB",
                "insert_point",
                self.open_insert_point,
            ),
            (
                "used_part_component_button",
                "Used Part Component",
                "Collect part component dari Excel program",
                "used_part_component",
                self.open_used_part_component,
            ),
            (
                "component_usage_button",
                "Component Usage Finder",
                "Cari component dipakai di model dan PCB apa saja",
                "component_usage",
                self.open_component_usage,
            ),
            (
                "common_feeder_reuse_button",
                "Common Feeder Reuse",
                "Cek substitute component aman atau conflict",
                "feeder_reuse",
                self.open_common_feeder_reuse,
            ),
            (
                "model_feeder_group_button",
                "Model Fix Feeder Groups",
                "Kelompokkan PCB yang component usage-nya mirip",
                "model_group",
                self.open_model_feeder_group,
            ),
        ]
        for index, (attr_name, title, subtitle, icon_name, callback) in enumerate(tools):
            button = self._create_tool_button(title, subtitle, icon_name, callback)
            setattr(self, attr_name, button)
            row, column = divmod(index, 4)
            tools_layout.addWidget(button, row, column)

        menu_card.layout.addLayout(tools_layout)
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

    def _build_feeder_mapping_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.back_feeder_mapping_button = QPushButton("Back to Tools")
        self.back_feeder_mapping_button.clicked.connect(self.open_menu)
        title = QLabel("Feeder Mapping Generator")
        title.setObjectName("TitleLabel")
        top.addWidget(self.back_feeder_mapping_button)
        top.addWidget(title)
        top.addStretch(1)
        layout.addLayout(top)

        self.feeder_mapping_page = FeederMappingPage(self.thread_pool, self.theme_manager)
        layout.addWidget(self.feeder_mapping_page, 1)
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

    def _build_used_part_component_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.back_used_part_component_button = QPushButton("Back to Tools")
        self.back_used_part_component_button.clicked.connect(self.open_menu)
        title = QLabel("PCBA Model Used Part Component Collector")
        title.setObjectName("TitleLabel")
        top.addWidget(self.back_used_part_component_button)
        top.addWidget(title)
        top.addStretch(1)
        layout.addLayout(top)

        self.used_part_component_page = UsedPartComponentPage(self.thread_pool, self.theme_manager)
        layout.addWidget(self.used_part_component_page, 1)
        return page

    def _build_component_usage_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.back_component_usage_button = QPushButton("Back to Tools")
        self.back_component_usage_button.clicked.connect(self.open_menu)
        title = QLabel("Component Usage Finder")
        title.setObjectName("TitleLabel")
        top.addWidget(self.back_component_usage_button)
        top.addWidget(title)
        top.addStretch(1)
        layout.addLayout(top)

        self.component_usage_page = ComponentUsageFinderPage(self.thread_pool, self.theme_manager)
        layout.addWidget(self.component_usage_page, 1)
        return page

    def _build_common_feeder_reuse_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.back_common_feeder_reuse_button = QPushButton("Back to Tools")
        self.back_common_feeder_reuse_button.clicked.connect(self.open_menu)
        title = QLabel("Common Parts / Feeder Reuse Analyzer")
        title.setObjectName("TitleLabel")
        top.addWidget(self.back_common_feeder_reuse_button)
        top.addWidget(title)
        top.addStretch(1)
        layout.addLayout(top)

        self.common_feeder_reuse_page = CommonFeederReusePage(self.thread_pool, self.theme_manager)
        layout.addWidget(self.common_feeder_reuse_page, 1)
        return page

    def _build_model_feeder_group_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.back_model_feeder_group_button = QPushButton("Back to Tools")
        self.back_model_feeder_group_button.clicked.connect(self.open_menu)
        title = QLabel("Model Fix Feeder Group Analyzer")
        title.setObjectName("TitleLabel")
        top.addWidget(self.back_model_feeder_group_button)
        top.addWidget(title)
        top.addStretch(1)
        layout.addLayout(top)

        self.model_feeder_group_page = ModelFeederGroupPage(self.thread_pool, self.theme_manager)
        layout.addWidget(self.model_feeder_group_page, 1)
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
        self.stack.setCurrentIndex(4)

    def open_new_pcb(self):
        self.stack.setCurrentIndex(5)

    def open_insert_point(self):
        self.stack.setCurrentIndex(6)

    def open_used_part_component(self):
        self.stack.setCurrentIndex(7)

    def open_component_usage(self):
        self.stack.setCurrentIndex(8)

    def open_common_feeder_reuse(self):
        self.stack.setCurrentIndex(9)

    def open_model_feeder_group(self):
        self.stack.setCurrentIndex(10)

    def open_worksheet_bom(self):
        self.stack.setCurrentIndex(2)

    def open_feeder_mapping(self):
        self.stack.setCurrentIndex(3)

    def _refresh_icons(self):
        theme = self.theme_manager.theme
        for button in self.tool_buttons:
            button.setIcon(make_icon(button._tool_icon_name, theme["accent"]))
        if hasattr(self, "back_button"):
            self.back_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_worksheet_bom_button"):
            self.back_worksheet_bom_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_feeder_mapping_button"):
            self.back_feeder_mapping_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_all_in_one_button"):
            self.back_all_in_one_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_new_pcb_button"):
            self.back_new_pcb_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_insert_point_button"):
            self.back_insert_point_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_used_part_component_button"):
            self.back_used_part_component_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_component_usage_button"):
            self.back_component_usage_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_common_feeder_reuse_button"):
            self.back_common_feeder_reuse_button.setIcon(make_icon("arrow_left", theme["text"]))
        if hasattr(self, "back_model_feeder_group_button"):
            self.back_model_feeder_group_button.setIcon(make_icon("arrow_left", theme["text"]))
