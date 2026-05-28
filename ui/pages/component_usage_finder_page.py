from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
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
from services.component_usage_finder_service import (
    ComponentUsageFinderConfig,
    export_component_usage_result,
    find_component_usage,
    suggest_export_name,
)
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu
from workers.task_runner import TaskWorker


class ComponentUsageFinderPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.search_result = None
        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Component Usage Finder")
        title.setObjectName("SectionTitle")
        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        source_card = Card()
        source_title = QLabel("Source Data")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)

        input_row = QHBoxLayout()
        input_label = QLabel("Component P/N:")
        input_label.setMinimumWidth(150)
        self.component_input = QLineEdit()
        self.component_input.setPlaceholderText("Masukkan component part number")
        self.component_input.returnPressed.connect(self.search_usage)
        input_row.addWidget(input_label)
        input_row.addWidget(self.component_input, 1)
        source_card.layout.addLayout(input_row)

        self.folder_picker = FilePicker("Folder Induk PCB:")
        self.folder_picker.browse_requested.connect(self.browse_source_folder)
        source_card.layout.addWidget(self.folder_picker)
        root.addWidget(source_card)

        action_bar = QHBoxLayout()
        self.search_btn = QPushButton("Search")
        self.search_btn.setObjectName("PrimaryButton")
        self.search_btn.clicked.connect(self.search_usage)
        self.copy_btn = QPushButton("Copy Results")
        self.copy_btn.clicked.connect(self.copy_results)
        self.export_btn = QPushButton("Export Excel")
        self.export_btn.setObjectName("SuccessButton")
        self.export_btn.clicked.connect(self.export_results)
        self.clear_btn = QPushButton("Clear Results")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_results)
        action_bar.addWidget(self.search_btn)
        action_bar.addWidget(self.copy_btn)
        action_bar.addWidget(self.export_btn)
        action_bar.addWidget(self.clear_btn)
        action_bar.addStretch(1)
        root.addLayout(action_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        table_card = Card()
        table_title = QLabel("Preview Results")
        table_title.setObjectName("SectionTitle")
        table_card.layout.addWidget(table_title)
        self.result_model = RecordTableModel(
            [
                ColumnSpec("no", "No", Qt.AlignCenter, 70),
                ColumnSpec("model_part_number", "Model Part Number", Qt.AlignLeft, 280),
                ColumnSpec("pcb_part_number", "PCB Part Number", Qt.AlignLeft, 240),
            ],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.result_model)
        self.result_table = QTableView()
        configure_table(self.result_table, self.result_model)
        install_copy_menu(self.result_table, self.result_model)
        table_card.layout.addWidget(self.result_table, 1)
        root.addWidget(table_card, 1)

        log_card = Card()
        log_title = QLabel("Scan Log")
        log_title.setObjectName("SectionTitle")
        log_card.layout.addWidget(log_title)
        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumHeight(110)
        self.log_output.setPlaceholderText("Skipped file dan detail source akan muncul di sini.")
        log_card.layout.addWidget(self.log_output)
        root.addWidget(log_card)

        self.register_busy_widgets(
            self.search_btn,
            self.copy_btn,
            self.export_btn,
            self.clear_btn,
            self.component_input,
            self.folder_picker.button,
        )

    def browse_source_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Pilih Folder Induk PCB")
        if folder_path:
            self.folder_picker.set_path(folder_path)

    def search_usage(self):
        if not self._validate_before_search():
            return

        config = ComponentUsageFinderConfig(
            component_part_number=self.component_input.text(),
            source_folder=self.folder_picker.path(),
        )
        self.result_model.set_records([])
        self.log_output.clear()
        self.search_result = None

        worker = None

        def task():
            return find_component_usage(
                config,
                progress_callback=lambda percent, message: worker.signals.progress.emit(percent, message),
            )

        worker = TaskWorker(task)
        worker._busy_text = "Scanning folders..."
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, "Scanning folders..."))
        worker.signals.progress.connect(self._on_search_progress)
        worker.signals.result.connect(self._on_search_done)
        worker.signals.error.connect(self._show_worker_error)
        worker.signals.finished.connect(lambda w=worker: self._finish_worker(w))
        self.thread_pool.start(worker)

    def _on_search_progress(self, percent, message):
        self.status_label.setText(message)
        self.progress.setVisible(True)
        self.progress.setRange(0, 100)
        self.progress.setValue(max(0, min(100, percent)))

    def _on_search_done(self, result):
        self.search_result = result
        records = [
            {
                "no": index,
                "model_part_number": row.model_part_number,
                "pcb_part_number": row.pcb_part_number,
            }
            for index, row in enumerate(result.rows, start=1)
        ]
        self.result_model.set_records(records)
        self.log_output.setPlainText(self._build_log_text(result))

        if result.total_files == 0:
            self.status_label.setText("Tidak ada file Excel ditemukan")
            QMessageBox.information(self, "Data kosong", "Tidak ada file Excel program di folder yang dipilih.")
        elif records:
            self.status_label.setText(f"Search complete: {len(records)} result(s) found")
        else:
            self.status_label.setText("No result found")

    def copy_results(self):
        if not self.result_model.records:
            QMessageBox.information(self, "Copy Results", "Tidak ada hasil untuk dicopy.")
            return
        QApplication.clipboard().setText(self.result_model.all_as_tsv(include_headers=True))
        self.status_label.setText("Results copied")

    def export_results(self):
        if self.search_result is None:
            QMessageBox.information(self, "Export Excel", "Belum ada hasil pencarian untuk diexport.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Component Usage Result",
            suggest_export_name(self.search_result.component_part_number),
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        try:
            saved_path = export_component_usage_result(self.search_result, output_path)
        except Exception as exc:
            QMessageBox.warning(self, "Export gagal", str(exc))
            return

        self.status_label.setText(f"Exported: {Path(saved_path).name}")
        self.status_label.setToolTip(saved_path)
        QMessageBox.information(self, "Export Excel", f"File berhasil dibuat:\n{saved_path}")

    def clear_results(self):
        self.search_result = None
        self.result_model.set_records([])
        self.log_output.clear()
        self.status_label.setText("Ready")
        self.status_label.setToolTip("")

    def _validate_before_search(self):
        component_part_number = self.component_input.text().strip()
        folder_path = self.folder_picker.path()

        if not component_part_number:
            QMessageBox.warning(self, "Input belum lengkap", "Component Part Number belum diisi.")
            return False
        if not folder_path:
            QMessageBox.warning(self, "Input belum lengkap", "Folder Induk PCB belum dipilih.")
            return False
        if not Path(folder_path).is_dir():
            QMessageBox.warning(self, "Folder tidak ditemukan", f"Folder Induk PCB tidak ditemukan:\n{folder_path}")
            return False
        return True

    def _build_log_text(self, result):
        lines = [
            f"Component: {result.component_part_number}",
            f"Excel files found: {result.total_files}",
            f"Files read: {result.read_files}",
            f"Skipped/error files: {len(result.skipped_files)}",
        ]

        if result.matched_files:
            lines.append("")
            lines.append("Matched source files:")
            for match in result.matched_files:
                rows_text = ", ".join(str(row) for row in match.found_rows) or "-"
                models_text = ", ".join(match.model_part_numbers) or "-"
                pcb_text = match.pcb_part_number
                if match.revision and pcb_text != "-":
                    pcb_text = f"{pcb_text}({match.revision})"
                lines.append(
                    f"- {match.source_folder} | {match.source_file} | PCB: {pcb_text} | Row: {rows_text} | Model: {models_text}"
                )

        if result.skipped_files:
            lines.append("")
            lines.append("Skipped/error:")
            for skipped in result.skipped_files[:80]:
                lines.append(f"- {skipped}")
            if len(result.skipped_files) > 80:
                lines.append(f"- ... {len(result.skipped_files) - 80} more")

        return "\n".join(lines)
