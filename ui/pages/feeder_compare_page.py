from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTableView,
    QVBoxLayout,
)

from models.table_model import ColumnSpec, RecordTableModel
from services.feeder_compare_service import (
    STATUS_ADD,
    STATUS_CNG,
    STATUS_DEL,
    STATUS_FILTER_ALL,
    STATUS_MOVE,
    STATUS_OPTIONS,
    compare_feeder_files,
    export_feeder_compare_result,
    suggest_export_name,
)
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.status_badge import StatusBadge
from widgets.table_tools import configure_table, install_copy_menu


class FeederComparePage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.compare_result = None
        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Feeder Compare")
        title.setObjectName("SectionTitle")
        self.summary_label = QLabel("0 DIFF")
        self.summary_label.setObjectName("MutedLabel")
        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addWidget(self.summary_label)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        source_card = Card()
        source_title = QLabel("Source Files")
        source_title.setObjectName("SectionTitle")
        
        mode_layout = QHBoxLayout()
        mode_label = QLabel("Mode:")
        self.mode_combo = QComboBox()
        self.mode_combo.addItems([
            "NPM vs NPM",
            "NPM vs NPM Feeder TXT",
            "CM602 vs CM602",
            "CM602 vs CM602 Feeder TXT"
        ])
        mode_layout.addWidget(mode_label)
        mode_layout.addWidget(self.mode_combo, 1)
        
        source_card.layout.addWidget(source_title)
        source_card.layout.addLayout(mode_layout)
        self.old_picker = FilePicker("Program A / Reference:")
        self.old_picker.browse_requested.connect(self.browse_old_file)
        self.new_picker = FilePicker("Program B / Target:")
        self.new_picker.browse_requested.connect(self.browse_new_file)
        source_card.layout.addWidget(self.old_picker)
        source_card.layout.addWidget(self.new_picker)
        root.addWidget(source_card)

        action_bar = QHBoxLayout()
        self.compare_btn = QPushButton("Compare Feeders")
        self.compare_btn.setObjectName("PrimaryButton")
        self.compare_btn.clicked.connect(self.compare_feed_data)
        self.copy_btn = QPushButton("Copy Results")
        self.copy_btn.clicked.connect(self.copy_results)
        self.export_btn = QPushButton("Export Excel")
        self.export_btn.setObjectName("SuccessButton")
        self.export_btn.clicked.connect(self.export_results)
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_all)
        self.copy_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        action_bar.addWidget(self.compare_btn)
        action_bar.addWidget(self.copy_btn)
        action_bar.addWidget(self.export_btn)
        action_bar.addStretch(1)
        action_bar.addWidget(self.clear_btn)
        root.addLayout(action_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        result_card = Card()
        result_header = QHBoxLayout()
        table_title = QLabel("Comparison Results")
        table_title.setObjectName("SectionTitle")
        self.add_badge = StatusBadge("0 ADD", STATUS_ADD)
        self.move_badge = StatusBadge("0 MOVE", STATUS_MOVE)
        self.cng_badge = StatusBadge("0 CNG", STATUS_CNG)
        self.del_badge = StatusBadge("0 DEL", STATUS_DEL)
        filter_label = QLabel("Status:")
        filter_label.setObjectName("MutedLabel")
        self.status_filter = QComboBox()
        self.status_filter.addItems(STATUS_OPTIONS)
        self.status_filter.setCurrentText(STATUS_FILTER_ALL)
        self.status_filter.currentTextChanged.connect(self.apply_filters)
        search_label = QLabel("Search:")
        search_label.setObjectName("MutedLabel")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search location, part number, position, atau description")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self.apply_filters)
        result_header.addWidget(table_title)
        result_header.addWidget(self.add_badge)
        result_header.addWidget(self.move_badge)
        result_header.addWidget(self.cng_badge)
        result_header.addWidget(self.del_badge)
        result_header.addStretch(1)
        result_header.addWidget(filter_label)
        result_header.addWidget(self.status_filter)
        result_header.addWidget(search_label)
        result_header.addWidget(self.search_input, 1)
        result_card.layout.addLayout(result_header)

        self.result_model = RecordTableModel(
            [
                ColumnSpec("no", "No", Qt.AlignCenter, 60),
                ColumnSpec("status", "Status", Qt.AlignCenter, 85),
                ColumnSpec("old_location", "Program A Location", Qt.AlignCenter, 140),
                ColumnSpec("old_part_number", "Program A Part Number", Qt.AlignLeft, 220),
                ColumnSpec("old_table", "A Table", Qt.AlignCenter, 80),
                ColumnSpec("old_slot", "A Slot", Qt.AlignCenter, 80),
                ColumnSpec("old_position", "A Position", Qt.AlignCenter, 140),
                ColumnSpec("new_location", "Program B Location", Qt.AlignCenter, 140),
                ColumnSpec("new_part_number", "Program B Part Number", Qt.AlignLeft, 220),
                ColumnSpec("new_table", "B Table", Qt.AlignCenter, 80),
                ColumnSpec("new_slot", "B Slot", Qt.AlignCenter, 80),
                ColumnSpec("new_position", "B Position", Qt.AlignCenter, 140),
                ColumnSpec("description", "Description", Qt.AlignLeft, 360),
            ],
            status_key="status",
            theme=self.theme_manager.theme,
        )
        self.register_model(self.result_model)
        self.result_table = QTableView()
        configure_table(self.result_table, self.result_model, wrap_headers=True)
        install_copy_menu(self.result_table, self.result_model)
        result_card.layout.addWidget(self.result_table, 1)
        root.addWidget(result_card, 1)

        self.register_busy_widgets(
            self.compare_btn,
            self.clear_btn,
            self.status_filter,
            self.search_input,
            self.old_picker.button,
            self.new_picker.button,
        )

    def _browse_file(self, title):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            title,
            "",
            "All Supported Files (*.txt *.crb);;Text Files (*.txt);;NPM Export (*.crb);;All Files (*)",
        )
        return file_path

    def browse_old_file(self):
        file_path = self._browse_file("Select Program A / Reference")
        if file_path:
            self.old_picker.set_path(file_path)
            self._clear_compare_output()
            self.status_label.setText("Program A selected")

    def browse_new_file(self):
        file_path = self._browse_file("Select Program B / Target")
        if file_path:
            self.new_picker.set_path(file_path)
            self._clear_compare_output()
            self.status_label.setText("Program B selected")

    def compare_feed_data(self):
        if not self._validate_inputs():
            return

        old_path = self.old_picker.path()
        new_path = self.new_picker.path()
        
        mode = self.mode_combo.currentText()
        if mode == "NPM vs NPM":
            old_parser, new_parser = "NPM", "NPM"
        elif mode == "NPM vs NPM Feeder TXT":
            old_parser, new_parser = "NPM", "NPM"
        elif mode == "CM602 vs CM602":
            old_parser, new_parser = "CM602", "CM602"
        elif mode == "CM602 vs CM602 Feeder TXT":
            old_parser, new_parser = "CM602", "CM602"
        else:
            old_parser, new_parser = "NPM", "NPM"
            
        self.compare_result = None
        self.result_model.set_records([])
        self.copy_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.run_worker(
            lambda o=old_path, n=new_path, op=old_parser, np=new_parser: compare_feeder_files(o, n, op, np),
            self._on_compare_done,
            "Comparing feeders...",
        )

    def _on_compare_done(self, result):
        self.compare_result = result
        self.add_badge.set_value(f"{result.add_count} ADD", STATUS_ADD)
        self.move_badge.set_value(f"{result.move_count} MOVE", STATUS_MOVE)
        self.cng_badge.set_value(f"{result.cng_count} CNG", STATUS_CNG)
        self.del_badge.set_value(f"{result.del_count} DEL", STATUS_DEL)
        self.export_btn.setEnabled(True)
        self.apply_filters()

        if result.rows:
            self.copy_btn.setEnabled(True)
            self.status_label.setText(
                f"Done: {len(result.rows)} diff(s) | A {result.old_count} feeders | B {result.new_count} feeders"
            )
        else:
            self.copy_btn.setEnabled(False)
            self.status_label.setText("All feeder setup match")
            QMessageBox.information(self, "All Match", "Semua setup feeder sudah match.")

    def apply_filters(self, *_):
        rows = self.compare_result.rows if self.compare_result is not None else []
        status_filter = self.status_filter.currentText()
        query = self.search_input.text().strip().lower()
        tokens = [token for token in query.split() if token]

        filtered_rows = []
        for row in rows:
            if status_filter != STATUS_FILTER_ALL and row.get("status") != status_filter:
                continue
            if tokens and not all(token in self._search_text(row) for token in tokens):
                continue
            filtered_rows.append(row)

        records = [dict(row, no=index) for index, row in enumerate(filtered_rows, start=1)]
        self.result_model.set_records(records)
        self._update_summary(len(filtered_rows), len(rows))

    def copy_results(self):
        if not self.result_model.records:
            QMessageBox.information(self, "Copy Results", "Tidak ada hasil untuk dicopy.")
            return
        QApplication.clipboard().setText(self.result_model.all_as_tsv(include_headers=True))
        self.status_label.setText("Results copied")

    def export_results(self):
        if self.compare_result is None:
            QMessageBox.information(self, "Export Excel", "Belum ada hasil compare untuk diexport.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Feeder Compare Result",
            suggest_export_name(self.old_picker.path(), self.new_picker.path()),
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        try:
            saved_path = export_feeder_compare_result(self.compare_result, output_path)
        except Exception as exc:
            QMessageBox.warning(self, "Export gagal", str(exc))
            return

        self.status_label.setText(f"Exported: {Path(saved_path).name}")
        self.status_label.setToolTip(saved_path)
        QMessageBox.information(self, "Export Excel", f"File berhasil dibuat:\n{saved_path}")

    def clear_all(self):
        self.old_picker.clear()
        self.new_picker.clear()
        self.search_input.clear()
        self.status_filter.setCurrentText(STATUS_FILTER_ALL)
        self._clear_compare_output()
        self.status_label.setText("Ready")
        self.status_label.setToolTip("")

    def _clear_compare_output(self):
        self.compare_result = None
        self.result_model.set_records([])
        self.copy_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.add_badge.set_value("0 ADD", STATUS_ADD)
        self.move_badge.set_value("0 MOVE", STATUS_MOVE)
        self.cng_badge.set_value("0 CNG", STATUS_CNG)
        self.del_badge.set_value("0 DEL", STATUS_DEL)
        self.summary_label.setText("0 DIFF")

    def _validate_inputs(self):
        old_path = self.old_picker.path()
        new_path = self.new_picker.path()

        if not old_path:
            QMessageBox.warning(self, "Input belum lengkap", "Program A / Reference belum dipilih.")
            return False
        if not Path(old_path).is_file():
            QMessageBox.warning(self, "File tidak ditemukan", f"Program A tidak ditemukan:\n{old_path}")
            return False
        if not new_path:
            QMessageBox.warning(self, "Input belum lengkap", "Program B / Target belum dipilih.")
            return False
        if not Path(new_path).is_file():
            QMessageBox.warning(self, "File tidak ditemukan", f"Program B tidak ditemukan:\n{new_path}")
            return False
        return True

    def _search_text(self, row):
        return " ".join(str(value) for value in row.values()).lower()

    def _update_summary(self, visible_count, total_count):
        if self.compare_result is None:
            self.summary_label.setText("0 DIFF")
            return

        summary = f"{total_count} DIFF | A {self.compare_result.old_count} FEEDERS | B {self.compare_result.new_count} FEEDERS"
        if visible_count != total_count:
            summary = f"{visible_count}/{summary}"
        self.summary_label.setText(summary)
