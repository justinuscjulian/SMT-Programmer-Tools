from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableView,
    QVBoxLayout,
)

from models.table_model import ColumnSpec, RecordTableModel
from services import all_in_one_service
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.status_badge import StatusBadge
from widgets.table_tools import configure_table, install_copy_menu


class AllInOneComparatorPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.all_results = []
        self.pickers = {}
        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("All In One Comparator")
        title.setObjectName("SectionTitle")
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedLabel")
        self.total_badge = StatusBadge("0 TOTAL", "INFO")
        self.ng_badge = StatusBadge("0 NG", "DEL")
        self.match_badge = StatusBadge("0 MATCH", "ADD")
        header.addWidget(title)
        header.addWidget(self.total_badge)
        header.addWidget(self.ng_badge)
        header.addWidget(self.match_badge)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        top_actions = QHBoxLayout()
        self.auto_import_btn = QPushButton("Auto-Import Semua File")
        self.auto_import_btn.setObjectName("PrimaryButton")
        self.auto_import_btn.clicked.connect(self.bulk_import)
        self.compare_btn = QPushButton("Start Compare")
        self.compare_btn.setObjectName("SuccessButton")
        self.compare_btn.clicked.connect(self.process_compare)
        self.clear_btn = QPushButton("Clear All")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_data)
        top_actions.addWidget(self.auto_import_btn)
        top_actions.addWidget(self.compare_btn)
        top_actions.addStretch(1)
        top_actions.addWidget(self.clear_btn)
        root.addLayout(top_actions)

        input_splitter = QSplitter(Qt.Horizontal)
        input_splitter.setChildrenCollapsible(False)
        input_splitter.addWidget(self._build_source_card())
        input_splitter.addWidget(self._build_target_card())
        input_splitter.setSizes([1, 1])
        root.addWidget(input_splitter)

        filter_bar = QHBoxLayout()
        filter_label = QLabel("Filter Status:")
        filter_label.setObjectName("SectionTitle")
        self.status_filter = QComboBox()
        self.status_filter.addItems(all_in_one_service.FILTER_OPTIONS)
        self.status_filter.setCurrentText(all_in_one_service.FILTER_NG_ONLY)
        self.status_filter.currentTextChanged.connect(self.apply_filter)
        filter_bar.addWidget(filter_label)
        filter_bar.addWidget(self.status_filter)
        filter_bar.addStretch(1)
        root.addLayout(filter_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        table_card = Card()
        table_title = QLabel("Summary Results")
        table_title.setObjectName("SectionTitle")
        table_card.layout.addWidget(table_title)
        self.result_model = RecordTableModel(
            [
                ColumnSpec("ref", "Ref", Qt.AlignCenter, 110),
                ColumnSpec("system", "System", Qt.AlignCenter, 80),
                ColumnSpec("status", "Status", Qt.AlignCenter, 180),
                ColumnSpec("source", "Source (Mesin/Ori)", Qt.AlignLeft, 380),
                ColumnSpec("target", "Target (TXT)", Qt.AlignLeft, 380),
            ],
            status_key="status",
            theme=self.theme_manager.theme,
        )
        self.register_model(self.result_model)
        self.result_table = QTableView()
        configure_table(self.result_table, self.result_model)
        install_copy_menu(self.result_table, self.result_model)
        table_card.layout.addWidget(self.result_table, 1)
        root.addWidget(table_card, 1)

        self.register_busy_widgets(
            self.auto_import_btn,
            self.compare_btn,
            self.clear_btn,
            *(picker.button for picker in self.pickers.values()),
        )

    def _build_source_card(self):
        card = Card()
        title = QLabel("File Mesin / Source (Kiri)")
        title.setObjectName("SectionTitle")
        card.layout.addWidget(title)
        self._add_picker(card, "npm_crb", "NPM (.crb)")
        self._add_picker(card, "cm602_machine", "CM602 (Machine)")
        self._add_picker(card, "bm_pos", "BM (.pos)")
        self._add_picker(card, "bom_ori", "BOM (.tsv/.csv)")
        return card

    def _build_target_card(self):
        card = Card()
        title = QLabel("File TXT / Target (Kanan)")
        title.setObjectName("SectionTitle")
        card.layout.addWidget(title)
        self._add_picker(card, "npm_txt", "NPM (.txt)")
        self._add_picker(card, "cm602_txt", "CM602 (.txt)")
        self._add_picker(card, "bm_txt", "BM (.txt)")
        self._add_picker(card, "bom_txt", "BOM (.txt)")
        return card

    def _add_picker(self, card, key, label):
        picker = FilePicker(label)
        picker.browse_requested.connect(lambda k=key: self.browse_file(k))
        self.pickers[key] = picker
        card.layout.addWidget(picker)

    def browse_file(self, key):
        title = f"Select {self.pickers[key].label.text()}"
        file_path, _ = QFileDialog.getOpenFileName(self, title, "", "All Files (*)")
        if file_path:
            self.pickers[key].set_path(file_path)

    def bulk_import(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Pilih Semua File", "", "All Files (*)")
        if not files:
            return
        paths = all_in_one_service.classify_bulk_files(files)
        for key, path in paths.items():
            if path and key in self.pickers:
                self.pickers[key].set_path(path)
        self.status_label.setText("Auto-import selesai")

    def process_compare(self):
        paths = {key: picker.path() for key, picker in self.pickers.items()}
        self.run_worker(lambda: all_in_one_service.process_compare(paths), self._on_compare_done, "Running all-in-one compare...")

    def _on_compare_done(self, results):
        self.all_results = results
        self.apply_filter()
        total_ng = sum(1 for result in results if result["status"] != "MATCH")
        total_match = sum(1 for result in results if result["status"] == "MATCH")
        self.total_badge.set_value(f"{len(results)} TOTAL", "INFO")
        self.ng_badge.set_value(f"{total_ng} NG", "DEL")
        self.match_badge.set_value(f"{total_match} MATCH", "ADD")

        if not results:
            QMessageBox.information(self, "Info", "Pilih file dulu, Gan!")
            self.status_label.setText("No file pair selected")
            return

        if total_ng == 0:
            QMessageBox.information(self, "Hasil Compare", "OK KOMPARE!")
        self.status_label.setText("Done")

    def apply_filter(self, *_):
        filter_type = self.status_filter.currentText()
        filtered = all_in_one_service.filter_results(self.all_results, filter_type)
        self.result_model.set_records(filtered)

    def clear_data(self):
        for picker in self.pickers.values():
            picker.clear()
        self.all_results = []
        self.result_model.set_records([])
        self.total_badge.set_value("0 TOTAL", "INFO")
        self.ng_badge.set_value("0 NG", "DEL")
        self.match_badge.set_value("0 MATCH", "ADD")
        self.status_filter.setCurrentText(all_in_one_service.FILTER_NG_ONLY)
        self.status_label.setText("")
