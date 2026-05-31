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
    QSpinBox,
    QTableView,
    QVBoxLayout,
)

from models.table_model import ColumnSpec, RecordTableModel
from services.model_feeder_group_service import (
    STATUS_GROUPED,
    STATUS_OPTIONS,
    ModelFeederGroupConfig,
    analyze_model_feeder_groups,
    export_model_feeder_group_result,
    suggest_export_name,
)
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu
from workers.task_runner import TaskWorker


class ModelFeederGroupPage(WorkerPage):
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
        title = QLabel("Model Fix Feeder Group Analyzer")
        title.setObjectName("SectionTitle")
        self.summary_label = QLabel("0 GROUPS")
        self.summary_label.setObjectName("MutedLabel")
        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addWidget(self.summary_label)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        source_card = Card()
        source_title = QLabel("Source Data")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)

        self.folder_picker = FilePicker("Folder Induk PCB:")
        self.folder_picker.browse_requested.connect(self.browse_source_folder)
        source_card.layout.addWidget(self.folder_picker)

        option_row = QHBoxLayout()
        similarity_label = QLabel("Min Similarity (%):")
        similarity_label.setMinimumWidth(150)
        self.similarity_spin = QSpinBox()
        self.similarity_spin.setRange(1, 100)
        self.similarity_spin.setValue(70)
        self.similarity_spin.setSuffix("%")
        shared_label = QLabel("Min Shared Parts:")
        shared_label.setMinimumWidth(130)
        self.shared_spin = QSpinBox()
        self.shared_spin.setRange(1, 9999)
        self.shared_spin.setValue(20)
        option_row.addWidget(similarity_label)
        option_row.addWidget(self.similarity_spin)
        option_row.addSpacing(16)
        option_row.addWidget(shared_label)
        option_row.addWidget(self.shared_spin)
        option_row.addStretch(1)
        source_card.layout.addLayout(option_row)
        root.addWidget(source_card)

        action_bar = QHBoxLayout()
        filter_label = QLabel("Status:")
        filter_label.setObjectName("MutedLabel")
        self.status_filter = QComboBox()
        self.status_filter.addItems(STATUS_OPTIONS)
        self.status_filter.setCurrentText(STATUS_GROUPED)
        self.status_filter.currentTextChanged.connect(self.apply_filters)
        self.analyze_btn = QPushButton("Analyze Groups")
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
        table_title = QLabel("Recommended Fix Feeder Groups")
        table_title.setObjectName("SectionTitle")
        search_label = QLabel("Search:")
        search_label.setObjectName("MutedLabel")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search group, PCB, model, atau component")
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
                ColumnSpec("group_id", "Group", Qt.AlignCenter, 80),
                ColumnSpec("status", "Status", Qt.AlignCenter, 95),
                ColumnSpec("member_count", "PCB Count", Qt.AlignCenter, 90),
                ColumnSpec("avg_similarity_percent", "Avg Similarity %", Qt.AlignCenter, 125),
                ColumnSpec("min_similarity_percent", "Min Similarity %", Qt.AlignCenter, 125),
                ColumnSpec("common_component_count", "Common Parts", Qt.AlignCenter, 110),
                ColumnSpec("union_component_count", "Union Parts", Qt.AlignCenter, 100),
                ColumnSpec("members", "PCB / Model Members", Qt.AlignLeft, 360),
                ColumnSpec("common_components", "Recommended Fixed Feeder Parts", Qt.AlignLeft, 420),
            ],
            status_key="status",
            theme=self.theme_manager.theme,
        )
        self.register_model(self.group_model)
        self.group_table = QTableView()
        configure_table(self.group_table, self.group_model, wrap_headers=True)
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
            self.similarity_spin,
            self.shared_spin,
        )

    def browse_source_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Pilih Folder Induk PCB")
        if folder_path:
            self.folder_picker.set_path(folder_path)

    def analyze_groups(self):
        if not self._validate_before_analysis():
            return

        config = ModelFeederGroupConfig(
            source_folder=self.folder_picker.path(),
            min_similarity_percent=self.similarity_spin.value(),
            min_shared_components=self.shared_spin.value(),
        )
        self.analysis_result = None
        self.group_model.set_records([])
        self.copy_btn.setEnabled(False)
        self.export_btn.setEnabled(False)

        worker = None

        def task():
            return analyze_model_feeder_groups(
                config,
                progress_callback=lambda percent, message: worker.signals.progress.emit(percent, message),
            )

        worker = TaskWorker(task)
        worker._busy_text = "Analyzing model groups..."
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, "Analyzing model groups..."))
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
            self.status_label.setText("Tidak ada BOM valid")
            QMessageBox.information(self, "Data kosong", 'Tidak ada sheet "BOM" valid yang bisa dianalisa.')
        else:
            self.status_label.setText(
                f"Done: {result.group_count} groups, {result.single_count} single, {result.model_count} models"
            )

    def apply_filters(self, *_):
        rows = self.analysis_result.group_rows if self.analysis_result is not None else []
        status_filter = self.status_filter.currentText()
        query = self.search_input.text().strip().lower()
        tokens = [token for token in query.split() if token]

        filtered_rows = []
        for row in rows:
            if status_filter != STATUS_OPTIONS[0] and row.get("status") != status_filter:
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
            "Save Model Fix Feeder Groups",
            suggest_export_name(),
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        try:
            saved_path = export_model_feeder_group_result(self.analysis_result, output_path)
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
        self.summary_label.setText("0 GROUPS")
        self.status_label.setText("Ready")
        self.status_label.setToolTip("")

    def _validate_before_analysis(self):
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
            self.summary_label.setText("0 GROUPS")
            return

        summary = (
            f"{total_count} ROWS | "
            f"{self.analysis_result.group_count} GROUPS | "
            f"{self.analysis_result.single_count} SINGLE"
        )
        if visible_count != total_count:
            summary = f"{visible_count}/{summary}"
        self.summary_label.setText(summary)
