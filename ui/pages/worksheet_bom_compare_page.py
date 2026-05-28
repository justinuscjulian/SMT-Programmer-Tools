from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
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
from services import worksheet_bom_service
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.status_badge import StatusBadge
from widgets.table_tools import configure_table, install_copy_menu


class WorksheetBomComparePage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.worksheet_summary = None
        self.bom_summary = None
        self.diff_results = []
        self.worksheet_path = ""
        self.bom_path = ""
        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Worksheet vs BOM Comparator")
        title.setObjectName("SectionTitle")
        self.add_badge = StatusBadge("0 ADD", "ADD")
        self.cng_badge = StatusBadge("0 CNG", "CNG")
        self.del_badge = StatusBadge("0 DEL", "DEL")
        self.status_label = QLabel("BOM scope: TOP SMT")
        self.status_label.setObjectName("MutedLabel")
        self.audit_time = QLabel("Last run: --:--:--")
        self.audit_time.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addWidget(self.add_badge)
        header.addWidget(self.cng_badge)
        header.addWidget(self.del_badge)
        header.addStretch(1)
        header.addWidget(self.status_label)
        header.addWidget(self.audit_time)
        root.addLayout(header)

        action_bar = QHBoxLayout()
        action_bar.setSpacing(10)
        self.compare_btn = QPushButton("Run Comparison")
        self.compare_btn.setObjectName("PrimaryButton")
        self.compare_btn.clicked.connect(self.compare_data)
        self.export_btn = QPushButton("Export Results")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_results)
        self.clear_btn = QPushButton("Clear All")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_all)
        action_bar.addWidget(self.compare_btn)
        action_bar.addWidget(self.export_btn)
        action_bar.addStretch(1)
        action_bar.addWidget(self.clear_btn)
        root.addLayout(action_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)

        upper = QSplitter(Qt.Horizontal)
        upper.setChildrenCollapsible(False)
        upper.addWidget(self._build_worksheet_card())
        upper.addWidget(self._build_bom_card())
        upper.setSizes([1, 1])

        splitter.addWidget(upper)
        splitter.addWidget(self._build_results_card())
        splitter.setSizes([500, 340])
        root.addWidget(splitter, 1)

        self.register_busy_widgets(
            self.compare_btn,
            self.clear_btn,
            self.worksheet_browse_btn,
            self.bom_browse_btn,
        )

    def _build_worksheet_card(self):
        card = Card()
        header = QHBoxLayout()
        title = QLabel("Worksheet File (.xlsx)")
        title.setObjectName("SectionTitle")
        self.worksheet_count = QLabel("0 PARTS")
        self.worksheet_count.setObjectName("MutedLabel")
        self.worksheet_file_label = QLabel("No file selected")
        self.worksheet_file_label.setObjectName("MutedLabel")
        self.worksheet_browse_btn = QPushButton("Browse")
        self.worksheet_browse_btn.clicked.connect(self.load_worksheet)
        header.addWidget(title)
        header.addWidget(self.worksheet_file_label, 1)
        header.addWidget(self.worksheet_count)
        header.addWidget(self.worksheet_browse_btn)
        card.layout.addLayout(header)

        self.worksheet_model = RecordTableModel(
            [
                ColumnSpec("PartNo", "Part Number", Qt.AlignCenter, 150),
                ColumnSpec("WorksheetQty", "CNT Total", Qt.AlignCenter, 90),
                ColumnSpec("FeedIds", "Feed ID", Qt.AlignLeft, 170),
                ColumnSpec("FeedSlots", "TBL/Feed", Qt.AlignLeft, 170),
                ColumnSpec("Spec", "Spec", Qt.AlignLeft, 320),
            ],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.worksheet_model)
        self.worksheet_table = QTableView()
        configure_table(self.worksheet_table, self.worksheet_model, wrap_headers=True)
        install_copy_menu(self.worksheet_table, self.worksheet_model)
        card.layout.addWidget(self.worksheet_table, 1)
        return card

    def _build_bom_card(self):
        card = Card()
        header = QHBoxLayout()
        title = QLabel("BOM File (.tsv)")
        title.setObjectName("SectionTitle")
        self.bom_count = QLabel("0 PARTS")
        self.bom_count.setObjectName("MutedLabel")
        self.bom_file_label = QLabel("No file selected")
        self.bom_file_label.setObjectName("MutedLabel")
        self.bom_browse_btn = QPushButton("Browse")
        self.bom_browse_btn.clicked.connect(self.load_bom)
        header.addWidget(title)
        header.addWidget(self.bom_file_label, 1)
        header.addWidget(self.bom_count)
        header.addWidget(self.bom_browse_btn)
        card.layout.addLayout(header)

        self.bom_model = RecordTableModel(
            [
                ColumnSpec("PartNo", "Part Number", Qt.AlignCenter, 150),
                ColumnSpec("BomQty", "BOM Qty", Qt.AlignCenter, 90),
                ColumnSpec("Side", "Side", Qt.AlignCenter, 70),
                ColumnSpec("RefDes", "RefDes", Qt.AlignLeft, 260),
                ColumnSpec("Spec", "Spec", Qt.AlignLeft, 320),
            ],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.bom_model)
        self.bom_table = QTableView()
        configure_table(self.bom_table, self.bom_model, wrap_headers=True)
        install_copy_menu(self.bom_table, self.bom_model)
        card.layout.addWidget(self.bom_table, 1)
        return card

    def _build_results_card(self):
        card = Card()
        header = QHBoxLayout()
        title = QLabel("Comparison Results")
        title.setObjectName("SectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        card.layout.addLayout(header)

        self.result_model = RecordTableModel(
            [
                ColumnSpec("no", "No.", Qt.AlignCenter, 55),
                ColumnSpec("part_no", "Part Number", Qt.AlignCenter, 150),
                ColumnSpec("worksheet_qty", "Worksheet CNT", Qt.AlignCenter, 110),
                ColumnSpec("bom_qty", "BOM Qty", Qt.AlignCenter, 90),
                ColumnSpec("delta", "Delta", Qt.AlignCenter, 80),
                ColumnSpec("feed_ids", "Feed ID", Qt.AlignLeft, 160),
                ColumnSpec("refdes", "BOM RefDes", Qt.AlignLeft, 220),
                ColumnSpec("type", "Type", Qt.AlignCenter, 80),
                ColumnSpec("desc", "Audit Description", Qt.AlignLeft, 320),
            ],
            status_key="type",
            theme=self.theme_manager.theme,
        )
        self.register_model(self.result_model)
        self.result_table = QTableView()
        configure_table(self.result_table, self.result_model, wrap_headers=True)
        install_copy_menu(self.result_table, self.result_model, clean_copy=True)
        card.layout.addWidget(self.result_table, 1)
        return card

    def load_worksheet(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Worksheet File",
            "",
            "Worksheet (*.xlsx *.xlsm);;Excel (*.xlsx *.xlsm *.xls);;All Files (*)",
        )
        if not file_path:
            return
        self.run_worker(
            lambda path=file_path: worksheet_bom_service.load_worksheet_summary(path),
            lambda result, path=file_path: self._on_worksheet_loaded(path, result),
            "Loading worksheet...",
        )

    def _on_worksheet_loaded(self, file_path, result):
        self.worksheet_path = file_path
        self.worksheet_summary = result
        self.worksheet_model.set_records(result.dataframe.to_dict("records"))
        skipped_text = f" | {result.skipped_rows} skipped" if result.skipped_rows else ""
        self.worksheet_count.setText(f"{result.row_count} PARTS | {result.total_qty} CNT{skipped_text}")
        self.worksheet_file_label.setText(Path(file_path).name)
        self.worksheet_file_label.setToolTip(file_path)
        self.status_label.setText("Worksheet loaded")

    def load_bom(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select BOM File",
            "",
            "BOM (*.tsv *.xlsx *.xls);;All Files (*)",
        )
        if not file_path:
            return
        self.run_worker(
            lambda path=file_path: worksheet_bom_service.load_bom_top_summary(path),
            lambda result, path=file_path: self._on_bom_loaded(path, result),
            "Loading BOM TOP data...",
        )

    def _on_bom_loaded(self, file_path, result):
        self.bom_path = file_path
        self.bom_summary = result
        self.bom_model.set_records(result.dataframe.to_dict("records"))
        assy_text = f" | Assy: {result.chassis_pn}" if result.chassis_pn else ""
        self.bom_count.setText(f"{result.row_count} PARTS | {result.total_qty} QTY{assy_text}")
        self.bom_file_label.setText(Path(file_path).name)
        self.bom_file_label.setToolTip(file_path)
        self.status_label.setText("BOM TOP loaded")

    def compare_data(self):
        if self.worksheet_summary is None or self.bom_summary is None:
            QMessageBox.warning(self, "Warning", "Import Worksheet dan BOM dulu!")
            return
        self.run_worker(
            lambda: worksheet_bom_service.compare_worksheet_bom(
                self.worksheet_summary.dataframe,
                self.bom_summary.dataframe,
            ),
            self._on_compare_done,
            "Running worksheet vs BOM compare...",
        )

    def _on_compare_done(self, diff_results):
        self.diff_results = diff_results
        add_count = sum(1 for item in diff_results if item["Type"] == "ADD")
        cng_count = sum(1 for item in diff_results if item["Type"] == "CNG")
        del_count = sum(1 for item in diff_results if item["Type"] == "DEL")
        self.add_badge.set_value(f"{add_count} ADD", "ADD")
        self.cng_badge.set_value(f"{cng_count} CNG", "CNG")
        self.del_badge.set_value(f"{del_count} DEL", "DEL")
        self.audit_time.setText(f"Last run: {datetime.now().strftime('%H:%M:%S')} Local")

        if not diff_results:
            records = [
                {
                    "no": "",
                    "part_no": "All Data Match!",
                    "worksheet_qty": "",
                    "bom_qty": "",
                    "delta": "",
                    "feed_ids": "",
                    "refdes": "",
                    "type": "MATCH",
                    "desc": "",
                }
            ]
            self.export_btn.setEnabled(False)
            QMessageBox.information(self, "All Data Match!", "Worksheet dan BOM TOP sudah match!")
        else:
            records = [
                {
                    "no": index,
                    "part_no": item["PartNo"],
                    "worksheet_qty": item["WorksheetQty"],
                    "bom_qty": item["BomQty"],
                    "delta": item["Delta"],
                    "feed_ids": item["FeedIds"],
                    "refdes": item["RefDes"],
                    "type": item["Type"],
                    "desc": item["Description"],
                    "_diff_keys": self._diff_keys(item["Type"]),
                }
                for index, item in enumerate(diff_results, 1)
            ]
            self.export_btn.setEnabled(True)

        self.result_model.set_records(records)
        self.status_label.setText("Done")

    def _diff_keys(self, diff_type):
        if diff_type == "CNG":
            return ["worksheet_qty", "bom_qty", "delta"]
        if diff_type == "ADD":
            return ["part_no", "bom_qty", "refdes"]
        if diff_type == "DEL":
            return ["part_no", "worksheet_qty", "feed_ids"]
        return []

    def export_results(self):
        if not self.diff_results:
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "Export Worksheet vs BOM Results", "", "Excel (*.xlsx)")
        if not file_path:
            return
        self.run_worker(
            lambda path=file_path: worksheet_bom_service.export_worksheet_bom_results(self.diff_results, path),
            lambda output_path: QMessageBox.information(self, "Success", f"Exported to:\n{output_path}"),
            "Exporting results...",
        )

    def clear_all(self):
        self.worksheet_summary = None
        self.bom_summary = None
        self.diff_results = []
        self.worksheet_path = ""
        self.bom_path = ""
        self.worksheet_model.set_records([])
        self.bom_model.set_records([])
        self.result_model.set_records([])
        self.worksheet_count.setText("0 PARTS")
        self.bom_count.setText("0 PARTS")
        self.worksheet_file_label.setText("No file selected")
        self.worksheet_file_label.setToolTip("")
        self.bom_file_label.setText("No file selected")
        self.bom_file_label.setToolTip("")
        self.add_badge.set_value("0 ADD", "ADD")
        self.cng_badge.set_value("0 CNG", "CNG")
        self.del_badge.set_value("0 DEL", "DEL")
        self.audit_time.setText("Last run: --:--:--")
        self.export_btn.setEnabled(False)
        self.status_label.setText("BOM scope: TOP SMT")
