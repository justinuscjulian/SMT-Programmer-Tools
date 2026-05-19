import os
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QFileDialog,
    QFrame,
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
from services import history_service, machine_service
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.status_badge import StatusBadge
from widgets.table_tools import configure_table, install_copy_menu


class MachineComparePage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.machine_type = "NPM"
        self.machine_df = None
        self.program_df = None
        self.diff_results = []
        self.machine_diff_rows = []
        self.program_diff_rows = []
        self.machine_file = ""
        self.program_file = ""
        self.last_machine_file_path = ""

        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        action_bar = QHBoxLayout()
        action_bar.setSpacing(10)
        self.compare_btn = QPushButton("Run Machine Audit")
        self.compare_btn.setObjectName("PrimaryButton")
        self.compare_btn.clicked.connect(self.compare_data)
        self.export_btn = QPushButton("Export Results")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_results)
        self.sync_btn = QPushButton("Sync to .POS (Beta)")
        self.sync_btn.setObjectName("SuccessButton")
        self.sync_btn.setVisible(False)
        self.sync_btn.setEnabled(False)
        self.sync_btn.clicked.connect(self.sync_pos_file)
        self.clear_btn = QPushButton("Clear All")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_all)
        action_bar.addWidget(self.compare_btn)
        action_bar.addWidget(self.export_btn)
        action_bar.addWidget(self.sync_btn)
        action_bar.addStretch(1)
        action_bar.addWidget(self.clear_btn)
        root.addLayout(action_bar)

        type_bar = QHBoxLayout()
        label = QLabel("Machine Type Selection:")
        label.setObjectName("SectionTitle")
        type_bar.addWidget(label)
        self.type_group = QButtonGroup(self)
        self.type_group.setExclusive(True)
        for machine_type in ["NPM", "CM602", "BM221"]:
            button = QPushButton(machine_type)
            button.setObjectName("SegmentedButton")
            button.setCheckable(True)
            button.setChecked(machine_type == self.machine_type)
            self.type_group.addButton(button)
            type_bar.addWidget(button)
        self.type_group.buttonClicked.connect(lambda button: self.on_machine_type_change(button.text()))
        type_bar.addStretch(1)
        root.addLayout(type_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        upper = QSplitter(Qt.Horizontal)
        upper.setChildrenCollapsible(False)
        upper.addWidget(self._build_machine_card())
        upper.addWidget(self._build_program_card())
        upper.setSizes([1, 1])
        splitter.addWidget(upper)
        splitter.addWidget(self._build_results_card())
        splitter.setSizes([520, 340])
        root.addWidget(splitter, 1)

        self.register_busy_widgets(self.compare_btn, self.clear_btn, self.machine_browse_btn, self.program_browse_btn)

    def _build_machine_card(self):
        card = Card()
        header = QHBoxLayout()
        self.machine_title = QLabel("Machine File (.crb)")
        self.machine_title.setObjectName("SectionTitle")
        self.machine_file_label = QLabel("No file selected")
        self.machine_file_label.setObjectName("MutedLabel")
        self.machine_count = QLabel("0 ROWS")
        self.machine_count.setObjectName("MutedLabel")
        self.machine_browse_btn = QPushButton("Browse")
        self.machine_browse_btn.clicked.connect(self.load_machine_file)
        header.addWidget(self.machine_title)
        header.addWidget(self.machine_file_label, 1)
        header.addWidget(self.machine_count)
        header.addWidget(self.machine_browse_btn)
        card.layout.addLayout(header)

        self.machine_model = RecordTableModel(
            [
                ColumnSpec("circuit", "Circuit No", Qt.AlignCenter, 120),
                ColumnSpec("x", "X Coordinate", Qt.AlignCenter, 120),
                ColumnSpec("y", "Y Coordinate", Qt.AlignCenter, 120),
                ColumnSpec("angle", "Angle", Qt.AlignCenter, 90),
                ColumnSpec("partno", "Parts Number", Qt.AlignCenter, 130),
                ColumnSpec("parts", "Parts", Qt.AlignLeft, 180),
            ],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.machine_model)
        self.machine_table = QTableView()
        configure_table(self.machine_table, self.machine_model)
        install_copy_menu(self.machine_table, self.machine_model)
        card.layout.addWidget(self.machine_table, 1)
        return card

    def _build_program_card(self):
        card = Card()
        header = QHBoxLayout()
        title = QLabel("Program File (.txt)")
        title.setObjectName("SectionTitle")
        self.program_file_label = QLabel("No file selected")
        self.program_file_label.setObjectName("MutedLabel")
        self.program_count = QLabel("0 ROWS")
        self.program_count.setObjectName("MutedLabel")
        self.program_browse_btn = QPushButton("Browse")
        self.program_browse_btn.clicked.connect(self.load_program_file)
        header.addWidget(title)
        header.addWidget(self.program_file_label, 1)
        header.addWidget(self.program_count)
        header.addWidget(self.program_browse_btn)
        card.layout.addLayout(header)

        self.program_model = RecordTableModel(
            [
                ColumnSpec("circuit", "Circuit No", Qt.AlignCenter, 120),
                ColumnSpec("x", "X Coordinate", Qt.AlignCenter, 120),
                ColumnSpec("y", "Y Coordinate", Qt.AlignCenter, 120),
                ColumnSpec("angle", "Angle", Qt.AlignCenter, 90),
                ColumnSpec("partno", "Parts Number", Qt.AlignCenter, 130),
                ColumnSpec("parts", "Parts", Qt.AlignLeft, 180),
            ],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.program_model)
        self.program_table = QTableView()
        configure_table(self.program_table, self.program_model)
        install_copy_menu(self.program_table, self.program_model)
        card.layout.addWidget(self.program_table, 1)
        return card

    def _build_results_card(self):
        card = Card()
        header = QHBoxLayout()
        title = QLabel("Machine Audit Results")
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

        result_splitter = QSplitter(Qt.Horizontal)
        result_splitter.setChildrenCollapsible(False)
        result_splitter.addWidget(self._build_result_side_panel("Machine Data", "machine"))
        result_splitter.addWidget(self._build_result_side_panel("Program File", "program"))
        result_splitter.setSizes([1, 1])
        card.layout.addWidget(result_splitter, 1)
        return card

    def _build_result_side_panel(self, title_text, panel_type):
        panel = QFrame()
        panel.setObjectName("SubPanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        title = QLabel(title_text)
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        columns = [
            ColumnSpec("circuit", "Circuit No", Qt.AlignCenter, 120),
            ColumnSpec("x", "X Coordinate", Qt.AlignCenter, 120),
            ColumnSpec("y", "Y Coordinate", Qt.AlignCenter, 120),
            ColumnSpec("angle", "Angle", Qt.AlignCenter, 90),
            ColumnSpec("partno", "Parts Number", Qt.AlignCenter, 130),
            ColumnSpec("parts", "Parts", Qt.AlignLeft, 180),
            ColumnSpec("type", "Type", Qt.AlignCenter, 80),
        ]
        model = RecordTableModel(columns, status_key="type", theme=self.theme_manager.theme)
        self.register_model(model)
        table = QTableView()
        configure_table(table, model)
        install_copy_menu(table, model, allow_cell_column=True)
        layout.addWidget(table, 1)

        if panel_type == "machine":
            self.machine_result_model = model
            self.machine_result_table = table
        else:
            self.program_result_model = model
            self.program_result_table = table

        return panel

    def on_machine_type_change(self, machine_type):
        self.machine_type = machine_type
        self.clear_all()
        self.sync_btn.setVisible(machine_type == "BM221")
        if machine_type == "NPM":
            self.machine_title.setText("Machine File (.crb)")
        elif machine_type == "BM221":
            self.machine_title.setText("Machine File (.POS)")
        else:
            self.machine_title.setText("Machine File (Any)")

    def load_machine_file(self):
        if self.machine_type == "NPM":
            file_filter = "Machine File (*.crb)"
        elif self.machine_type == "BM221":
            file_filter = "Machine File (*.POS)"
        else:
            file_filter = "All Files (*)"

        file_path, _ = QFileDialog.getOpenFileName(self, f"Select {self.machine_type} Machine File", "", file_filter)
        if not file_path:
            return

        self.run_worker(
            lambda path=file_path, mtype=self.machine_type: machine_service.load_machine_file(path, mtype),
            lambda df, path=file_path: self._on_machine_loaded(path, df),
            f"Loading {self.machine_type} machine file...",
        )

    def _on_machine_loaded(self, file_path, dataframe):
        self.machine_df = dataframe
        self.last_machine_file_path = file_path
        self.machine_file = os.path.basename(file_path)
        self.machine_model.set_records(dataframe.to_dict("records"))
        self.machine_count.setText(f"{len(dataframe)} ROWS")
        self.machine_file_label.setText(self.machine_file)
        self.status_label.setText("Machine file loaded")

    def load_program_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Program File (.txt)", "", "Program File (*.txt)")
        if not file_path:
            return
        self.run_worker(
            lambda path=file_path, mtype=self.machine_type: machine_service.load_program_file(path, mtype),
            lambda df, path=file_path: self._on_program_loaded(path, df),
            "Loading program file...",
        )

    def _on_program_loaded(self, file_path, dataframe):
        self.program_df = dataframe
        self.program_file = os.path.basename(file_path)
        self.program_model.set_records(dataframe.to_dict("records"))
        self.program_count.setText(f"{len(dataframe)} ROWS")
        self.program_file_label.setText(self.program_file)
        self.status_label.setText("Program file loaded")

    def compare_data(self):
        if self.machine_df is None or self.program_df is None:
            QMessageBox.warning(self, "Warning", "Import both files first!")
            return
        self.run_worker(lambda: machine_service.compare_machine(self.machine_df, self.program_df), self._on_compare_done, "Running machine audit...")

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
            self.machine_diff_rows = [{"circuit": "All Data Match!", "x": "", "y": "", "angle": "", "partno": "", "parts": "", "type": "MATCH"}]
            self.program_diff_rows = [{"circuit": "All Data Match!", "x": "", "y": "", "angle": "", "partno": "", "parts": "", "type": "MATCH"}]
            self.export_btn.setEnabled(False)
            self.sync_btn.setEnabled(False)
        else:
            self.machine_diff_rows, self.program_diff_rows = machine_service.build_machine_diff_preview(
                self.machine_df,
                self.program_df,
                diff_results,
            )
            self.export_btn.setEnabled(True)
            self.sync_btn.setEnabled(self.machine_type == "BM221")

        self.machine_result_model.set_records(self.machine_diff_rows)
        self.program_result_model.set_records(self.program_diff_rows)
        history_service.save_history(machine_service.machine_history_entry(self.machine_file, self.program_file, diff_results))
        self.status_label.setText("Done")
        if all_data_match:
            QMessageBox.information(self, "All Data Match!", "All Data Match!")

    def export_results(self):
        if not self.diff_results:
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "Export", "", "Excel (*.xlsx)")
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"
        self.run_worker(
            lambda path=file_path: machine_service.export_machine_preview(self.machine_diff_rows, self.program_diff_rows, path),
            lambda _: QMessageBox.information(self, "Success", f"Exported to:\n{file_path}"),
            "Exporting results...",
        )

    def sync_pos_file(self):
        if self.machine_type != "BM221" or not self.last_machine_file_path:
            return
        self.run_worker(lambda: machine_service.prepare_bm221_sync(self.last_machine_file_path, self.diff_results), self._on_sync_prepared, "Preparing POS sync...")

    def _on_sync_prepared(self, result):
        if result.skipped_tray_circuits:
            circuits = ", ".join(sorted(result.skipped_tray_circuits))
            QMessageBox.warning(
                self,
                "TRAY Component Detected",
                "Perbedaan pada slot TRAY (Z200+) terdeteksi pada circuit:\n"
                f"{circuits}\n\n"
                "Sistem HANYA MENGABAIKAN komponen TRAY. Harap ubah komponen TRAY secara manual di Notepad!",
            )

        if result.replacements_made == 0 and not result.skipped_tray_circuits:
            QMessageBox.information(self, "Info", "Tidak ada perubahan Parts (CNG) pada slot Feeder yang perlu di-sync.")
            return

        initial = f"SYNC_{os.path.basename(self.last_machine_file_path)}"
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Synced .POS File", initial, "POS File (*.POS)")
        if not save_path:
            return
        if not save_path.lower().endswith(".pos"):
            save_path += ".POS"

        self.run_worker(
            lambda path=save_path, content=result.content: machine_service.write_pos_file(content, path),
            lambda _, path=save_path: QMessageBox.information(self, "Success", f"File POS berhasil di-sync dan disimpan di:\n{path}"),
            "Saving synced POS file...",
        )

    def clear_all(self):
        self.machine_df = None
        self.program_df = None
        self.diff_results = []
        self.machine_diff_rows = []
        self.program_diff_rows = []
        self.machine_file = ""
        self.program_file = ""
        self.last_machine_file_path = ""
        self.machine_model.set_records([])
        self.program_model.set_records([])
        self.machine_result_model.set_records([])
        self.program_result_model.set_records([])
        self.machine_count.setText("0 ROWS")
        self.program_count.setText("0 ROWS")
        self.machine_file_label.setText("No file selected")
        self.program_file_label.setText("No file selected")
        self.add_badge.set_value("0 ADD", "ADD")
        self.cng_badge.set_value("0 CNG", "CNG")
        self.del_badge.set_value("0 DEL", "DEL")
        self.audit_time.setText("Last run: --:--:--")
        self.status_label.setText("")
        self.export_btn.setEnabled(False)
        self.sync_btn.setEnabled(False)
