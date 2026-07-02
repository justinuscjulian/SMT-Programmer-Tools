import os
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from models.table_model import ColumnSpec, RecordTableModel
from services import bom_service, history_service
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.status_badge import StatusBadge
from widgets.table_tools import configure_table, install_copy_menu


COMPARE_MODE_TXT_TO_BOM = "txt_to_bom"
COMPARE_MODE_TXT_TO_TXT = "txt_to_txt"


class BomComparePage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.compare_mode = COMPARE_MODE_TXT_TO_BOM
        self.reference_df = None
        self.raw_df = None
        self.diff_results = []
        self.reference_file = ""
        self.raw_file = ""
        self.raw_meta = None

        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

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

        mode_bar = QHBoxLayout()
        mode_bar.setSpacing(10)
        mode_label = QLabel("Comparison Mode:")
        mode_label.setObjectName("SectionTitle")
        self.txt_bom_mode_btn = QPushButton("TXT vs BOM File")
        self.txt_bom_mode_btn.setObjectName("SegmentedButton")
        self.txt_bom_mode_btn.setCheckable(True)
        self.txt_bom_mode_btn.setProperty("compare_mode", COMPARE_MODE_TXT_TO_BOM)
        self.txt_txt_mode_btn = QPushButton("TXT vs TXT")
        self.txt_txt_mode_btn.setObjectName("SegmentedButton")
        self.txt_txt_mode_btn.setCheckable(True)
        self.txt_txt_mode_btn.setProperty("compare_mode", COMPARE_MODE_TXT_TO_TXT)
        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)
        self.mode_group.addButton(self.txt_bom_mode_btn)
        self.mode_group.addButton(self.txt_txt_mode_btn)
        self.txt_bom_mode_btn.setChecked(True)
        self.mode_group.buttonClicked.connect(lambda button: self.on_compare_mode_change(button.property("compare_mode")))
        mode_bar.addWidget(mode_label)
        mode_bar.addWidget(self.txt_bom_mode_btn)
        mode_bar.addWidget(self.txt_txt_mode_btn)
        mode_bar.addStretch(1)
        root.addLayout(mode_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)

        upper = QSplitter(Qt.Horizontal)
        upper.setChildrenCollapsible(False)
        upper.addWidget(self._build_reference_card())
        upper.addWidget(self._build_raw_card())
        upper.setSizes([1, 1])

        splitter.addWidget(upper)
        splitter.addWidget(self._build_results_card())
        splitter.setSizes([520, 340])
        root.addWidget(splitter, 1)

        self.register_busy_widgets(
            self.compare_btn,
            self.clear_btn,
            self.ref_browse_btn,
            self.raw_browse_btn,
            self.txt_bom_mode_btn,
            self.txt_txt_mode_btn,
        )

    def _build_reference_card(self):
        card = Card()
        header = QHBoxLayout()
        self.ref_title = QLabel("BOM Excel Reference (.txt)")
        self.ref_title.setObjectName("SectionTitle")
        self.ref_count = QLabel("0 ROWS")
        self.ref_count.setObjectName("MutedLabel")
        self.ref_file_label = QLabel("No file selected")
        self.ref_file_label.setObjectName("MutedLabel")
        self.ref_browse_btn = QPushButton("Browse")
        self.ref_browse_btn.clicked.connect(self.load_reference)
        header.addWidget(self.ref_title)
        header.addWidget(self.ref_file_label, 1)
        header.addWidget(self.ref_count)
        header.addWidget(self.ref_browse_btn)
        card.layout.addLayout(header)

        self.reference_model = RecordTableModel(
            [
                ColumnSpec("Circuit", "Circuit No", Qt.AlignCenter, 130),
                ColumnSpec("PartNo", "Part Number", Qt.AlignCenter, 220),
            ],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.reference_model)
        self.reference_table = QTableView()
        configure_table(self.reference_table, self.reference_model)
        install_copy_menu(self.reference_table, self.reference_model)
        card.layout.addWidget(self.reference_table, 1)
        return card

    def _build_raw_card(self):
        card = Card()
        header = QHBoxLayout()
        self.raw_title = QLabel("BOM File (.tsv/.xlsx/.xls)")
        self.raw_title.setObjectName("SectionTitle")
        self.raw_count = QLabel("0 ROWS")
        self.raw_count.setObjectName("MutedLabel")
        self.raw_file_label = QLabel("No file selected")
        self.raw_file_label.setObjectName("MutedLabel")
        self.raw_browse_btn = QPushButton("Browse")
        self.raw_browse_btn.clicked.connect(self.load_raw_bom)
        header.addWidget(self.raw_title)
        header.addWidget(self.raw_file_label, 1)
        header.addWidget(self.raw_count)
        header.addWidget(self.raw_browse_btn)
        card.layout.addLayout(header)

        self.raw_bom_columns = [
            ColumnSpec("Chassis", "Date/Time", Qt.AlignCenter, 130),
            ColumnSpec("Circuit", "Circuit No", Qt.AlignCenter, 120),
            ColumnSpec("PartNo", "PCB PN", Qt.AlignCenter, 150),
            ColumnSpec("Spec", "Specification", Qt.AlignLeft, 240),
            ColumnSpec("Side", "Side", Qt.AlignCenter, 70),
        ]
        self.source_txt_columns = [
            ColumnSpec("Circuit", "Circuit No", Qt.AlignCenter, 130),
            ColumnSpec("PartNo", "Part Number", Qt.AlignCenter, 220),
        ]
        self.raw_model = RecordTableModel(self.raw_bom_columns, theme=self.theme_manager.theme)
        self.register_model(self.raw_model)
        self.raw_table = QTableView()
        configure_table(self.raw_table, self.raw_model)
        install_copy_menu(self.raw_table, self.raw_model)
        card.layout.addWidget(self.raw_table, 1)
        return card

    def _build_results_card(self):
        card = Card()
        header = QHBoxLayout()
        title = QLabel("Comparison Results")
        title.setObjectName("SectionTitle")
        self.add_badge = StatusBadge("0 ADD", "ADD")
        self.cng_badge = StatusBadge("0 CNG", "CNG")
        self.del_badge = StatusBadge("0 DEL", "DEL")
        self.status_label = QLabel("")
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
        card.layout.addLayout(header)

        self.result_model = RecordTableModel(
            self._result_columns(),
            status_key="type",
            theme=self.theme_manager.theme,
        )
        self.register_model(self.result_model)
        self.result_table = QTableView()
        configure_table(self.result_table, self.result_model)
        install_copy_menu(
            self.result_table, 
            self.result_model, 
            clean_copy=True, 
            copy_all_excluded_keys={"side"},
            copy_all_include_headers=False,
            copy_all_sort_keys=["txt_part", "tsv_part", "type"]
        )
        card.layout.addWidget(self.result_table, 1)
        return card

    def _result_columns(self):
        source_header = "Part No (Source .txt)" if self.compare_mode == COMPARE_MODE_TXT_TO_TXT else "Part No (Source)"
        return [
            ColumnSpec("no", "No.", Qt.AlignCenter, 60),
            ColumnSpec("circuit", "Circuit No", Qt.AlignCenter, 120),
            ColumnSpec("side", "Side", Qt.AlignCenter, 70),
            ColumnSpec("txt_part", "Part No (Reference .txt)", Qt.AlignCenter, 190),
            ColumnSpec("tsv_part", source_header, Qt.AlignCenter, 190),
            ColumnSpec("type", "Type", Qt.AlignCenter, 80),
            ColumnSpec("desc", "Audit Description", Qt.AlignLeft, 360),
        ]

    def _apply_table_columns(self, table, model, columns):
        model.set_columns(columns)
        for idx, column in enumerate(columns):
            table.setColumnWidth(idx, column.width)

    def on_compare_mode_change(self, compare_mode):
        if not compare_mode or compare_mode == self.compare_mode:
            return

        self.compare_mode = compare_mode
        self._reset_source()
        self._reset_results()
        self._sync_mode_ui()
        self.status_label.setText(f"Mode: {self._mode_label()}")

    def _sync_mode_ui(self):
        if self.compare_mode == COMPARE_MODE_TXT_TO_TXT:
            self.raw_title.setText("BOM Excel Source (.txt)")
            self._apply_table_columns(self.raw_table, self.raw_model, self.source_txt_columns)
        else:
            self.raw_title.setText("BOM File (.tsv/.xlsx/.xls)")
            self._apply_table_columns(self.raw_table, self.raw_model, self.raw_bom_columns)
        self._apply_table_columns(self.result_table, self.result_model, self._result_columns())

    def _reset_source(self):
        self.raw_df = None
        self.raw_file = ""
        self.raw_meta = None
        self.raw_model.set_records([])
        self.raw_count.setText("0 ROWS")
        self.raw_file_label.setText("No file selected")

    def _reset_results(self):
        self.diff_results = []
        self.result_model.set_records([])
        self.add_badge.set_value("0 ADD", "ADD")
        self.cng_badge.set_value("0 CNG", "CNG")
        self.del_badge.set_value("0 DEL", "DEL")
        self.audit_time.setText("Last run: --:--:--")
        self.export_btn.setEnabled(False)

    def _mode_label(self):
        if self.compare_mode == COMPARE_MODE_TXT_TO_TXT:
            return "TXT vs TXT"
        return "TXT vs BOM File"

    def _source_label(self):
        if self.compare_mode == COMPARE_MODE_TXT_TO_TXT:
            return "Source .txt"
        return "BOM File"

    def load_reference(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Reference .txt", "", "Text (*.txt)")
        if not file_path:
            return
        self.run_worker(
            lambda path=file_path: bom_service.load_reference_txt(path, check_duplicate_circuits=True),
            lambda df, path=file_path: self._on_reference_loaded(path, df),
            "Loading reference TXT...",
        )

    def _on_reference_loaded(self, file_path, dataframe):
        self.reference_df = dataframe
        self.reference_file = os.path.basename(file_path)
        self.reference_model.set_records(dataframe.to_dict("records"))
        self.ref_count.setText(f"{len(dataframe)} ROWS")
        self.ref_file_label.setText(self.reference_file)
        self._reset_results()
        self.status_label.setText("Reference loaded")

    def load_raw_bom(self):
        if self.compare_mode == COMPARE_MODE_TXT_TO_TXT:
            file_path, _ = QFileDialog.getOpenFileName(self, "Select Source .txt", "", "Text (*.txt)")
            if not file_path:
                return
            self.run_worker(
                lambda path=file_path: bom_service.load_reference_txt(path, check_duplicate_circuits=True),
                lambda df, path=file_path: self._on_source_txt_loaded(path, df),
                "Loading source TXT...",
            )
            return

        file_path, _ = QFileDialog.getOpenFileName(self, "Select Raw BOM", "", "BOM (*.tsv *.xlsx *.xls);;All Files (*)")
        if not file_path:
            return
        self.run_worker(
            lambda path=file_path: bom_service.load_raw_bom(path, check_duplicate_circuits=True),
            lambda result, path=file_path: self._on_raw_loaded(path, result),
            "Loading raw BOM...",
        )

    def _on_raw_loaded(self, file_path, result):
        self.raw_df = result.dataframe
        self.raw_file = os.path.basename(file_path)
        self.raw_meta = result
        columns = [
            ColumnSpec("Chassis", result.timestamp, Qt.AlignCenter, 130),
            ColumnSpec("Circuit", "Circuit No", Qt.AlignCenter, 120),
            ColumnSpec("PartNo", result.pcb_pn if result.pcb_pn else "Part No", Qt.AlignCenter, 150),
            ColumnSpec("Spec", "Specification", Qt.AlignLeft, 240),
            ColumnSpec("Side", "Side", Qt.AlignCenter, 70),
        ]
        self._apply_table_columns(self.raw_table, self.raw_model, columns)
        self.raw_model.set_records(result.dataframe.to_dict("records"))
        self.raw_count.setText(f"{len(result.dataframe)} ROWS | Assy: {result.chassis_pn}")
        self.raw_file_label.setText(self.raw_file)
        self._reset_results()
        self.status_label.setText("Raw BOM loaded")

    def _on_source_txt_loaded(self, file_path, dataframe):
        self.raw_df = dataframe
        self.raw_file = os.path.basename(file_path)
        self.raw_meta = None
        self._apply_table_columns(self.raw_table, self.raw_model, self.source_txt_columns)
        self.raw_model.set_records(dataframe.to_dict("records"))
        self.raw_count.setText(f"{len(dataframe)} ROWS")
        self.raw_file_label.setText(self.raw_file)
        self._reset_results()
        self.status_label.setText("Source TXT loaded")

    def compare_data(self):
        if self.reference_df is None or self.raw_df is None:
            QMessageBox.warning(self, "Warning", "Import both files first!")
            return
        self.run_worker(
            lambda: bom_service.compare_bom(
                self.reference_df,
                self.raw_df,
                reference_label="Reference .txt",
                source_label=self._source_label(),
            ),
            self._on_compare_done,
            "Running comparison...",
        )

    def _on_compare_done(self, diff_results):
        self.diff_results = diff_results
        add_count = sum(1 for item in diff_results if item[4] == "ADD")
        cng_count = sum(1 for item in diff_results if item[4] == "CNG")
        del_count = sum(1 for item in diff_results if item[4] == "DEL")
        self.add_badge.set_value(f"{add_count} ADD", "ADD")
        self.cng_badge.set_value(f"{cng_count} CNG", "CNG")
        self.del_badge.set_value(f"{del_count} DEL", "DEL")
        self.audit_time.setText(f"Last run: {datetime.now().strftime('%H:%M:%S')} Local")

        all_data_match = not diff_results
        if all_data_match:
            records = [{"no": "", "circuit": "All Data Match!", "side": "", "txt_part": "", "tsv_part": "", "type": "MATCH", "desc": ""}]
            self.export_btn.setEnabled(False)
        else:
            records = [
                {
                    "no": idx,
                    "circuit": circuit,
                    "side": side,
                    "txt_part": txt_part,
                    "tsv_part": tsv_part,
                    "type": diff_type,
                    "desc": desc,
                    "_diff_keys": self._bom_diff_keys(diff_type),
                }
                for idx, (circuit, side, txt_part, tsv_part, diff_type, desc) in enumerate(diff_results, 1)
            ]
            self.export_btn.setEnabled(True)

        self.result_model.set_records(records)
        history_service.save_history(
            {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "mode": self._mode_label(),
                "txt_file": self.reference_file,
                "tsv_file": self.raw_file,
                "add_count": add_count,
                "cng_count": cng_count,
                "del_count": del_count,
                "results": diff_results,
            }
        )
        self.status_label.setText("Done")
        if all_data_match:
            QMessageBox.information(self, "All Data Match!", "All Data Match!")

    def _bom_diff_keys(self, diff_type):
        if diff_type == "CNG":
            return ["txt_part", "tsv_part"]
        if diff_type == "ADD":
            return ["circuit", "side", "tsv_part"]
        if diff_type == "DEL":
            return ["circuit", "txt_part"]
        return []

    def export_results(self):
        if not self.diff_results:
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "Export", "", "Excel (*.xlsx)")
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"
        source_header = "Part (Source .txt)" if self.compare_mode == COMPARE_MODE_TXT_TO_TXT else "Part (Source)"
        self.run_worker(
            lambda path=file_path: bom_service.export_bom_results(
                self.diff_results,
                path,
                reference_header="Part (Reference .txt)",
                source_header=source_header,
            ),
            lambda _: QMessageBox.information(self, "Success", f"Exported to:\n{file_path}"),
            "Exporting results...",
        )

    def clear_all(self):
        self.reference_df = None
        self.reference_file = ""
        self.reference_model.set_records([])
        self.ref_count.setText("0 ROWS")
        self.ref_file_label.setText("No file selected")
        self._reset_source()
        self._reset_results()
        self._sync_mode_ui()
        self.status_label.setText("")
