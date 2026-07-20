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
    QRadioButton,
    QTableView,
    QTextEdit,
    QVBoxLayout,
)

from models.table_model import ColumnSpec, RecordTableModel
from services.table7_feeder_service import (
    Table7FeederConfig,
    analyze_table7_feeders,
    export_table7_result,
    suggest_table7_export_name,
)
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu
from workers.task_runner import TaskWorker


class Table7FeederPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.analysis_result = None
        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Table 7 Fix Feeder Generator")
        title.setObjectName("SectionTitle")
        self.summary_label = QLabel("0 PCB")
        self.summary_label.setObjectName("MutedLabel")
        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addWidget(self.summary_label)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        source_card = Card()
        source_title = QLabel("Source Data & Database")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)

        # File picker for Component Database
        self.db_picker = FilePicker("DB Table 7 (Excel):")
        self.db_picker.browse_requested.connect(self.browse_db_file)
        # Try to default if exists in current dir
        default_db = Path("KOMPONEN TABLE 7.xlsx").resolve()
        if default_db.exists():
            self.db_picker.set_path(str(default_db))
        source_card.layout.addWidget(self.db_picker)

        # File picker for Parent Folder
        self.folder_picker = FilePicker("Folder Induk PCB:")
        self.folder_picker.browse_requested.connect(self.browse_source_folder)
        source_card.layout.addWidget(self.folder_picker)

        # Mode Selection
        mode_layout = QHBoxLayout()
        self.mode_all = QRadioButton("Mode 1: Scan All in Folder")
        self.mode_all.setChecked(True)
        self.mode_list = QRadioButton("Mode 2: Filter by PCB List")
        self.mode_all.toggled.connect(self._toggle_mode)
        mode_layout.addWidget(self.mode_all)
        mode_layout.addWidget(self.mode_list)
        mode_layout.addStretch(1)
        source_card.layout.addLayout(mode_layout)

        self.list_input = QTextEdit()
        self.list_input.setPlaceholderText("Paste list PCB Part Number di sini (pisahkan dengan enter atau koma)...\nContoh:\nEAX67123456\nEAX67123457")
        self.list_input.setMaximumHeight(80)
        self.list_input.setVisible(False)
        source_card.layout.addWidget(self.list_input)

        root.addWidget(source_card)

        action_bar = QHBoxLayout()
        filter_label = QLabel("Status:")
        filter_label.setObjectName("MutedLabel")
        self.status_filter = QComboBox()
        self.status_filter.addItems(["SHOW ALL", "OK", "OVERLOAD (> 30)", "NO TABLE 7 PARTS"])
        self.status_filter.setCurrentText("SHOW ALL")
        self.status_filter.currentTextChanged.connect(self.apply_filters)
        
        self.analyze_btn = QPushButton("Generate Fix Feeder")
        self.analyze_btn.setObjectName("PrimaryButton")
        self.analyze_btn.clicked.connect(self.analyze_groups)
        
        self.copy_btn = QPushButton("Copy Results")
        self.copy_btn.clicked.connect(self.copy_results)
        
        self.export_btn = QPushButton("Export Excel")
        self.export_btn.setObjectName("SuccessButton")
        self.export_btn.clicked.connect(self.export_results)
        
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_results)
        self.copy_btn.setEnabled(False)
        self.export_btn.setEnabled(False)

        action_bar.addWidget(filter_label)
        action_bar.addWidget(self.status_filter)
        action_bar.addSpacing(10)
        action_bar.addWidget(self.analyze_btn)
        action_bar.addWidget(self.copy_btn)
        action_bar.addWidget(self.export_btn)
        action_bar.addStretch(1)
        action_bar.addWidget(self.clear_btn)
        root.addLayout(action_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        table_card = Card()
        table_header = QHBoxLayout()
        table_title = QLabel("Table 7 Fix Feeder Recommendations")
        table_title.setObjectName("SectionTitle")
        search_label = QLabel("Search:")
        search_label.setObjectName("MutedLabel")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search PCB or component")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self.apply_filters)
        table_header.addWidget(table_title)
        table_header.addStretch(1)
        table_header.addWidget(search_label)
        table_header.addWidget(self.search_input, 1)
        table_card.layout.addLayout(table_header)

        self.group_model = RecordTableModel(
            [
                ColumnSpec("no", "No", Qt.AlignCenter, 60),
                ColumnSpec("pcb_part_number", "PCB Part Number", Qt.AlignLeft, 180),
                ColumnSpec("status", "Status", Qt.AlignCenter, 130),
                ColumnSpec("table7_part_count", "Parts Used", Qt.AlignCenter, 90),
                ColumnSpec("slot_assignments", "Slot Assignments (1 - 30)", Qt.AlignLeft, 350),
                ColumnSpec("members", "PCB Members", Qt.AlignLeft, 250),
            ],
            status_key="status",
            theme=self.theme_manager.theme,
        )
        self.register_model(self.group_model)
        self.group_table = QTableView()
        configure_table(self.group_table, self.group_model, wrap_headers=True)
        # Enable multi-line for slot assignments if needed, but table_tools does it.
        install_copy_menu(self.group_table, self.group_model)
        table_card.layout.addWidget(self.group_table, 1)
        root.addWidget(table_card, 1)

        self.register_busy_widgets(
            self.analyze_btn,
            self.copy_btn,
            self.export_btn,
            self.clear_btn,
            self.status_filter,
            self.search_input,
            self.folder_picker.button,
            self.db_picker.button,
            self.mode_all,
            self.mode_list,
            self.list_input,
        )

    def _toggle_mode(self):
        self.list_input.setVisible(self.mode_list.isChecked())

    def browse_source_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Pilih Folder Induk PCB")
        if folder_path:
            self.folder_picker.set_path(folder_path)

    def browse_db_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Pilih File Database Table 7", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.db_picker.set_path(file_path)

    def analyze_groups(self):
        if not self._validate_before_analysis():
            return

        target_list = []
        if self.mode_list.isChecked():
            raw_text = self.list_input.toPlainText()
            target_list = [t.strip().upper() for t in raw_text.replace(",", "\n").split("\n") if t.strip()]
            if not target_list:
                QMessageBox.warning(self, "Input belum lengkap", "List PCB Part Number tidak boleh kosong di Mode 2.")
                return

        config = Table7FeederConfig(
            source_folder=self.folder_picker.path(),
            table7_ref_file=self.db_picker.path(),
            target_pcb_list=target_list,
        )
        self.analysis_result = None
        self.group_model.set_records([])
        self.copy_btn.setEnabled(False)
        self.export_btn.setEnabled(False)

        worker = None

        def task():
            return analyze_table7_feeders(
                config,
                progress_callback=lambda percent, message: worker.signals.progress.emit(percent, message),
            )

        worker = TaskWorker(task)
        worker._busy_text = "Analyzing Table 7 feeders..."
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, "Analyzing Table 7 feeders..."))
        worker.signals.progress.connect(self._on_analysis_progress)
        worker.signals.result.connect(self._on_analysis_done)
        worker.signals.error.connect(self._show_worker_error)
        worker.signals.finished.connect(lambda w=worker: self._finish_worker(w))
        self.thread_pool.start(worker)

    def _on_analysis_progress(self, percent, message):
        self.status_label.setText(message)
        self.progress.setVisible(True)
        self.progress.setRange(0, 100)
        self.progress.setValue(max(0, min(100, percent)))

    def _on_analysis_done(self, result):
        self.analysis_result = result
        self.copy_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
        self.apply_filters()

        if result.total_files == 0:
            self.status_label.setText("Tidak ada file Excel ditemukan")
            QMessageBox.information(self, "Data kosong", "Tidak ada file Excel program di folder yang dipilih.")
        elif result.model_count == 0:
            self.status_label.setText("Tidak ada PCB yang valid")
            QMessageBox.information(self, "Data kosong", 'Tidak ada PCB yang bisa dianalisa.')
        else:
            self.status_label.setText(f"Done: {result.model_count} PCB dianalisa")

    def apply_filters(self, *_):
        rows = self.analysis_result.pcb_rows if self.analysis_result is not None else []
        status_filter = self.status_filter.currentText()
        query = self.search_input.text().strip().lower()
        tokens = [token for token in query.split() if token]

        filtered_rows = []
        for row in rows:
            if status_filter != "SHOW ALL" and row.get("status") != status_filter:
                continue
            if tokens and not all(token in self._search_text(row) for token in tokens):
                continue
            filtered_rows.append(row)

        records = [dict(row, no=index) for index, row in enumerate(filtered_rows, start=1)]
        self.group_model.set_records(records)
        self._update_summary(len(filtered_rows), len(rows))

    def copy_results(self):
        if not self.group_model.records:
            QMessageBox.information(self, "Copy Results", "Tidak ada hasil untuk dicopy.")
            return
        QApplication.clipboard().setText(self.group_model.all_as_tsv(include_headers=True))
        self.status_label.setText("Results copied")

    def export_results(self):
        if self.analysis_result is None:
            QMessageBox.information(self, "Export Excel", "Belum ada hasil analisa untuk diexport.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Table 7 Fix Feeder Groups",
            suggest_table7_export_name(),
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        try:
            saved_path = export_table7_result(self.analysis_result, output_path)
        except Exception as exc:
            QMessageBox.warning(self, "Export gagal", str(exc))
            return

        self.status_label.setText(f"Exported: {Path(saved_path).name}")
        self.status_label.setToolTip(saved_path)
        QMessageBox.information(self, "Export Excel", f"File berhasil dibuat:\n{saved_path}")

    def clear_results(self):
        self.analysis_result = None
        self.folder_picker.clear()
        self.search_input.clear()
        self.group_model.set_records([])
        self.copy_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.summary_label.setText("0 PCB")
        self.status_label.setText("Ready")
        self.status_label.setToolTip("")

    def _validate_before_analysis(self):
        if not self.db_picker.path() or not Path(self.db_picker.path()).is_file():
            QMessageBox.warning(self, "Input belum lengkap", "File Database Komponen Table 7 belum dipilih atau tidak valid.")
            return False
            
        folder_path = self.folder_picker.path()
        if not folder_path:
            QMessageBox.warning(self, "Input belum lengkap", "Folder Induk PCB belum dipilih.")
            return False
        if not Path(folder_path).is_dir():
            QMessageBox.warning(self, "Folder tidak ditemukan", f"Folder Induk PCB tidak ditemukan:\n{folder_path}")
            return False
        return True

    def _search_text(self, row):
        return " ".join(str(value) for value in row.values()).lower()

    def _update_summary(self, visible_count, total_count):
        if self.analysis_result is None:
            self.summary_label.setText("0 PCB")
            return

        summary = f"{total_count} PCB TOTAL"
        if visible_count != total_count:
            summary = f"{visible_count}/{summary}"
        self.summary_label.setText(summary)
