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
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QTableView,
    QVBoxLayout,
)

from models.table_model import ColumnSpec, RecordTableModel
from services.common_feeder_reuse_service import (
    STATUS_OPTIONS,
    STATUS_SAFE,
    CommonFeederReuseConfig,
    analyze_common_feeder_reuse,
    export_common_feeder_reuse_result,
    suggest_export_name,
)
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu
from workers.task_runner import TaskWorker


class CommonFeederReusePage(WorkerPage):
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
        title = QLabel("Common Parts / Feeder Reuse Analyzer")
        title.setObjectName("SectionTitle")
        self.summary_label = QLabel("0 PAIRS")
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

        self.feeder_picker = FilePicker("Fixed Feeder Source:")
        self.feeder_picker.browse_requested.connect(self.browse_feeder_source)
        source_card.layout.addWidget(self.feeder_picker)

        candidate_label = QLabel("Candidate Component P/N:")
        candidate_label.setObjectName("MutedLabel")
        source_card.layout.addWidget(candidate_label)
        self.candidate_input = QPlainTextEdit()
        self.candidate_input.setMaximumHeight(92)
        self.candidate_input.setPlaceholderText(
            "Opsional. Isi satu atau beberapa component P/N yang mau dicek.\n"
            "Kosongkan untuk analisa semua component non-feeder dari folder PCB."
        )
        source_card.layout.addWidget(self.candidate_input)
        root.addWidget(source_card)

        action_bar = QHBoxLayout()
        filter_label = QLabel("Status:")
        filter_label.setObjectName("MutedLabel")
        self.status_filter = QComboBox()
        self.status_filter.addItems(STATUS_OPTIONS)
        self.status_filter.setCurrentText(STATUS_SAFE)
        self.status_filter.currentTextChanged.connect(self.apply_filters)
        self.analyze_btn = QPushButton("Analyze Reuse")
        self.analyze_btn.setObjectName("PrimaryButton")
        self.analyze_btn.clicked.connect(self.analyze_reuse)
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
        table_title = QLabel("Compatibility Results")
        table_title.setObjectName("SectionTitle")
        search_label = QLabel("Search:")
        search_label.setObjectName("MutedLabel")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search candidate, main P/N, slot, atau PCB conflict")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self.apply_filters)
        table_header.addWidget(table_title)
        table_header.addStretch(1)
        table_header.addWidget(search_label)
        table_header.addWidget(self.search_input, 1)
        table_card.layout.addLayout(table_header)

        self.result_model = RecordTableModel(
            [
                ColumnSpec("no", "No", Qt.AlignCenter, 60),
                ColumnSpec("status", "Status", Qt.AlignCenter, 95),
                ColumnSpec("candidate_part_number", "Candidate P/N", Qt.AlignLeft, 220),
                ColumnSpec("main_part_number", "Main Feeder P/N", Qt.AlignLeft, 220),
                ColumnSpec("location_code", "Location Code", Qt.AlignLeft, 140),
                ColumnSpec("table", "Table", Qt.AlignCenter, 70),
                ColumnSpec("slot", "Slot", Qt.AlignCenter, 70),
                ColumnSpec("position", "Position", Qt.AlignCenter, 130),
                ColumnSpec("candidate_usage_count", "Candidate Used In", Qt.AlignCenter, 120),
                ColumnSpec("main_usage_count", "Main Used In", Qt.AlignCenter, 105),
                ColumnSpec("conflict_count", "Conflict Count", Qt.AlignCenter, 110),
                ColumnSpec("conflict_programs", "Conflict PCB / Model", Qt.AlignLeft, 360),
            ],
            status_key="status",
            theme=self.theme_manager.theme,
        )
        self.register_model(self.result_model)
        self.result_table = QTableView()
        configure_table(self.result_table, self.result_model, wrap_headers=True)
        install_copy_menu(self.result_table, self.result_model)
        table_card.layout.addWidget(self.result_table, 1)
        root.addWidget(table_card, 1)

        log_card = Card()
        log_title = QLabel("Scan Log")
        log_title.setObjectName("SectionTitle")
        log_card.layout.addWidget(log_title)
        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumHeight(120)
        log_card.layout.addWidget(self.log_output)
        root.addWidget(log_card)

        self.register_busy_widgets(
            self.analyze_btn,
            self.copy_btn,
            self.export_btn,
            self.clear_btn,
            self.status_filter,
            self.search_input,
            self.candidate_input,
            self.folder_picker.button,
            self.feeder_picker.button,
        )

    def browse_source_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Pilih Folder Induk PCB")
        if folder_path:
            self.folder_picker.set_path(folder_path)

    def browse_feeder_source(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Fixed Feeder Source",
            "",
            "Feeder Source (*.txt *.xlsx *.xlsm *.xls);;NPM Export (*.txt);;Excel Workbook (*.xlsx *.xlsm *.xls);;All Files (*)",
        )
        if file_path:
            self.feeder_picker.set_path(file_path)

    def analyze_reuse(self):
        if not self._validate_before_analysis():
            return

        config = CommonFeederReuseConfig(
            source_folder=self.folder_picker.path(),
            feeder_source_path=self.feeder_picker.path(),
            candidate_part_numbers=self.candidate_input.toPlainText(),
        )
        self.analysis_result = None
        self.result_model.set_records([])
        self.log_output.clear()
        self.copy_btn.setEnabled(False)
        self.export_btn.setEnabled(False)

        worker = None

        def task():
            return analyze_common_feeder_reuse(
                config,
                progress_callback=lambda percent, message: worker.signals.progress.emit(percent, message),
            )

        worker = TaskWorker(task)
        worker._busy_text = "Analyzing feeder reuse..."
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, "Analyzing feeder reuse..."))
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
        self.log_output.setPlainText(self._build_log_text(result))
        self.apply_filters()

        if result.total_files == 0:
            self.status_label.setText("Tidak ada file Excel ditemukan")
            QMessageBox.information(self, "Data kosong", "Tidak ada file Excel program di folder yang dipilih.")
        else:
            self.status_label.setText(
                f"Done: {result.safe_count} SAFE, {result.conflict_count} CONFLICT, {result.check_count} CHECK"
            )

    def apply_filters(self, *_):
        rows = self.analysis_result.rows if self.analysis_result is not None else []
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
        self.result_model.set_records(records)
        self._update_summary(len(filtered_rows), len(rows))

    def copy_results(self):
        if not self.result_model.records:
            QMessageBox.information(self, "Copy Results", "Tidak ada hasil untuk dicopy.")
            return
        QApplication.clipboard().setText(self.result_model.all_as_tsv(include_headers=True))
        self.status_label.setText("Results copied")

    def export_results(self):
        if self.analysis_result is None:
            QMessageBox.information(self, "Export Excel", "Belum ada hasil analisa untuk diexport.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Common Feeder Reuse Result",
            suggest_export_name(),
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        try:
            saved_path = export_common_feeder_reuse_result(self.analysis_result, output_path)
        except Exception as exc:
            QMessageBox.warning(self, "Export gagal", str(exc))
            return

        self.status_label.setText(f"Exported: {Path(saved_path).name}")
        self.status_label.setToolTip(saved_path)
        QMessageBox.information(self, "Export Excel", f"File berhasil dibuat:\n{saved_path}")

    def clear_results(self):
        self.analysis_result = None
        self.folder_picker.clear()
        self.feeder_picker.clear()
        self.candidate_input.clear()
        self.search_input.clear()
        self.result_model.set_records([])
        self.log_output.clear()
        self.copy_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.summary_label.setText("0 PAIRS")
        self.status_label.setText("Ready")
        self.status_label.setToolTip("")

    def _validate_before_analysis(self):
        folder_path = self.folder_picker.path()
        feeder_path = self.feeder_picker.path()

        if not folder_path:
            QMessageBox.warning(self, "Input belum lengkap", "Folder Induk PCB belum dipilih.")
            return False
        if not Path(folder_path).is_dir():
            QMessageBox.warning(self, "Folder tidak ditemukan", f"Folder Induk PCB tidak ditemukan:\n{folder_path}")
            return False
        if not feeder_path:
            QMessageBox.warning(self, "Input belum lengkap", "Fixed Feeder Source belum dipilih.")
            return False
        if not Path(feeder_path).is_file():
            QMessageBox.warning(self, "File tidak ditemukan", f"Fixed Feeder Source tidak ditemukan:\n{feeder_path}")
            return False
        return True

    def _search_text(self, row):
        return " ".join(str(value) for value in row.values()).lower()

    def _update_summary(self, visible_count, total_count):
        if self.analysis_result is None:
            self.summary_label.setText("0 PAIRS")
            return

        summary = (
            f"{total_count} PAIRS | "
            f"{self.analysis_result.safe_count} SAFE | "
            f"{self.analysis_result.conflict_count} CONFLICT | "
            f"{self.analysis_result.check_count} CHECK"
        )
        if visible_count != total_count:
            summary = f"{visible_count}/{summary}"
        self.summary_label.setText(summary)

    def _build_log_text(self, result):
        lines = [
            f"Excel files found: {result.total_files}",
            f"Files read: {result.read_files}",
            f"Components found: {result.component_count}",
            f"Fixed feeder rows: {len(result.feeder_records)}",
            f"Candidate parts: {result.candidate_count}",
            f"SAFE pairs: {result.safe_count}",
            f"CONFLICT pairs: {result.conflict_count}",
            f"CHECK pairs: {result.check_count}",
            "",
            "Status rule:",
            "SAFE = candidate dan main feeder tidak pernah muncul bareng di PCB/model yang discan.",
            "CONFLICT = candidate dan main feeder muncul bareng minimal di satu PCB/model.",
            "CHECK = salah satu P/N tidak ditemukan di folder scan, jadi perlu confirm manual.",
        ]

        if result.skipped_files:
            lines.append("")
            lines.append("Skipped/error:")
            for skipped in result.skipped_files[:80]:
                lines.append(f"- {skipped}")
            if len(result.skipped_files) > 80:
                lines.append(f"- ... {len(result.skipped_files) - 80} more")

        return "\n".join(lines)
