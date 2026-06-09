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
from services.component_usage_plan_service import (
    ComponentUsagePlanConfig,
    export_component_usage_plan_result,
    find_component_usage_on_excel_plan,
    suggest_plan_export_name,
)
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu
from workers.task_runner import TaskWorker


class ComponentUsagePlanPage(WorkerPage):
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
        title = QLabel("Component Usage Finder on Excel Plan")
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

        self.plan_picker = FilePicker("Excel Plan File:")
        self.plan_picker.browse_requested.connect(self.browse_plan_file)
        source_card.layout.addWidget(self.plan_picker)

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
                ColumnSpec("no", "No", Qt.AlignCenter, 60),
                ColumnSpec("line", "LINE", Qt.AlignLeft, 150),
                ColumnSpec("wo_supply", "WO SUPPLY", Qt.AlignLeft, 210),
                ColumnSpec("dms_part_number", "DMS P/N", Qt.AlignLeft, 220),
                ColumnSpec("pcb_part_number", "PCB", Qt.AlignLeft, 220),
            ],
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
        self.log_output.setPlaceholderText("Detail parsing plan dan file yang diskip akan muncul di sini.")
        log_card.layout.addWidget(self.log_output)
        root.addWidget(log_card)

        self.register_busy_widgets(
            self.search_btn,
            self.copy_btn,
            self.export_btn,
            self.clear_btn,
            self.component_input,
            self.plan_picker.button,
            self.folder_picker.button,
        )

    def browse_plan_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Pilih Excel Plan",
            "",
            "Excel Workbook (*.xlsx *.xlsm *.xls)",
        )
        if file_path:
            self.plan_picker.set_path(file_path)

    def browse_source_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Pilih Folder Induk PCB")
        if folder_path:
            self.folder_picker.set_path(folder_path)

    def search_usage(self):
        if not self._validate_before_search():
            return

        config = ComponentUsagePlanConfig(
            component_part_number=self.component_input.text(),
            plan_file=self.plan_picker.path(),
            source_folder=self.folder_picker.path(),
        )
        self.result_model.set_records([])
        self.log_output.clear()
        self.search_result = None

        worker = None

        def task():
            return find_component_usage_on_excel_plan(
                config,
                progress_callback=lambda percent, message: worker.signals.progress.emit(percent, message),
            )

        worker = TaskWorker(task)
        worker._busy_text = "Scanning Excel plan..."
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, "Scanning Excel plan..."))
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
                "line": row.line,
                "wo_supply": row.wo_supply,
                "dms_part_number": row.dms_part_number,
                "pcb_part_number": row.pcb_part_number,
            }
            for index, row in enumerate(result.rows, start=1)
        ]
        self.result_model.set_records(records)
        self.log_output.setPlainText(self._build_log_text(result))

        if not result.plan_entries:
            self.status_label.setText("Tidak ada history valid di kolom V")
            QMessageBox.information(
                self,
                "Data kosong",
                "Tidak ada history program valid di kolom V yang PCB-nya match dengan kolom I.",
            )
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
            "Save Component Usage Plan Result",
            suggest_plan_export_name(self.search_result.component_part_number),
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        try:
            saved_path = export_component_usage_plan_result(self.search_result, output_path)
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
        plan_file = self.plan_picker.path()
        folder_path = self.folder_picker.path()

        if not component_part_number:
            QMessageBox.warning(self, "Input belum lengkap", "Component Part Number belum diisi.")
            return False
        if not plan_file:
            QMessageBox.warning(self, "Input belum lengkap", "File Excel plan belum dipilih.")
            return False
        if not Path(plan_file).is_file():
            QMessageBox.warning(self, "File tidak ditemukan", f"File Excel plan tidak ditemukan:\n{plan_file}")
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
            f"Plan file: {result.plan_file}",
            f"Folder induk: {result.source_folder}",
            f"History rows parsed and PCB matched: {len(result.plan_entries)}",
            f"Unique program targets: {result.unique_target_count}",
            f"PCB folders found: {result.pcb_folder_count}",
            f"Candidate program files: {result.candidate_file_count}",
            f"Program files read: {result.read_file_count}",
            f"Results: {len(result.rows)}",
            f"Skipped/error: {len(result.skipped_files)}",
        ]

        if result.rows:
            lines.append("")
            lines.append("Matched plan rows:")
            for row in result.rows[:80]:
                lines.append(
                    f"- {row.line} | {row.wo_supply} | {row.dms_part_number} | {row.pcb_part_number} | {row.sheet_name}!V{row.row_number}"
                )
            if len(result.rows) > 80:
                lines.append(f"- ... {len(result.rows) - 80} more")

        if result.matched_programs:
            lines.append("")
            lines.append("Matched program files:")
            for match in result.matched_programs[:80]:
                rows_text = ", ".join(str(value) for value in match.found_rows) or "-"
                lines.append(
                    f"- {match.main_part_number} | {match.pcb_part_number} | {match.source_folder} | {match.source_file} | BOM row: {rows_text}"
                )
            if len(result.matched_programs) > 80:
                lines.append(f"- ... {len(result.matched_programs) - 80} more")

        if result.skipped_files:
            lines.append("")
            lines.append("Skipped/error:")
            for skipped in result.skipped_files[:100]:
                lines.append(f"- {skipped}")
            if len(result.skipped_files) > 100:
                lines.append(f"- ... {len(result.skipped_files) - 100} more")

        return "\n".join(lines)

    def _join_short(self, values, limit=8):
        if not values:
            return ""
        shown = [str(value) for value in values[:limit]]
        if len(values) > limit:
            shown.append(f"+{len(values) - limit} more")
        return ", ".join(shown)
