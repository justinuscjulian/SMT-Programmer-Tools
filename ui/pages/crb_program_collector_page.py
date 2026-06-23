from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QTableView,
    QVBoxLayout,
)

from models.table_model import ColumnSpec, RecordTableModel
from services.crb_program_collector_service import (
    CrbProgramCollectorConfig,
    copy_crb_programs,
    export_crb_program_report,
    format_crb_program_report,
    scan_crb_programs,
    suggest_report_name,
)
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu
from workers.task_runner import TaskWorker


class CrbProgramCollectorPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.scan_result = None
        self.copy_result = None
        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("CRB Program Finder/Collector")
        title.setObjectName("SectionTitle")
        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        source_card = Card()
        source_title = QLabel("Source, Destination, and Filters")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)

        self.source_picker = FilePicker("Source Folder:")
        self.source_picker.browse_requested.connect(self.browse_source_folder)
        source_card.layout.addWidget(self.source_picker)

        self.destination_picker = FilePicker("Destination Folder:")
        self.destination_picker.browse_requested.connect(self.browse_destination_folder)
        source_card.layout.addWidget(self.destination_picker)

        part_label = QLabel("PCB Part Number List:")
        part_label.setObjectName("MutedLabel")
        self.part_input = QPlainTextEdit()
        self.part_input.setPlaceholderText("Satu PCB part number per baris, contoh:\nEAX01905902\nEAX01975401")
        self.part_input.setMinimumHeight(92)
        source_card.layout.addWidget(part_label)
        source_card.layout.addWidget(self.part_input)

        line_label = QLabel("Line Filter (optional):")
        line_label.setObjectName("MutedLabel")
        self.line_filter_input = QPlainTextEdit()
        self.line_filter_input.setPlaceholderText("Kosong = semua line. Contoh: INI3, INI4")
        self.line_filter_input.setMaximumHeight(62)
        source_card.layout.addWidget(line_label)
        source_card.layout.addWidget(self.line_filter_input)
        root.addWidget(source_card)

        action_bar = QHBoxLayout()
        self.scan_btn = QPushButton("Scan")
        self.scan_btn.setObjectName("PrimaryButton")
        self.scan_btn.clicked.connect(self.scan_programs)
        self.copy_btn = QPushButton("Copy")
        self.copy_btn.setObjectName("SuccessButton")
        self.copy_btn.clicked.connect(self.copy_matches)
        self.export_btn = QPushButton("Export Report")
        self.export_btn.clicked.connect(self.export_report)
        self.open_dest_btn = QPushButton("Open Destination Folder")
        self.open_dest_btn.clicked.connect(self.open_destination_folder)
        self.clear_btn = QPushButton("Clear Result")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_result)
        action_bar.addWidget(self.scan_btn)
        action_bar.addWidget(self.copy_btn)
        action_bar.addWidget(self.export_btn)
        action_bar.addWidget(self.open_dest_btn)
        action_bar.addStretch(1)
        action_bar.addWidget(self.clear_btn)
        root.addLayout(action_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        table_card = Card()
        table_title = QLabel("Matched CRB Files")
        table_title.setObjectName("SectionTitle")
        table_card.layout.addWidget(table_title)
        self.result_model = RecordTableModel(
            [
                ColumnSpec("no", "No", Qt.AlignCenter, 56),
                ColumnSpec("file_name", "File Name", Qt.AlignLeft, 280),
                ColumnSpec("matched_parts", "Matched PCB P/N", Qt.AlignLeft, 230),
                ColumnSpec("line", "Line", Qt.AlignCenter, 90),
                ColumnSpec("size_kb", "Size KB", Qt.AlignRight, 90),
                ColumnSpec("source_folder", "Source Folder", Qt.AlignLeft, 260),
                ColumnSpec("source_path", "Full Path", Qt.AlignLeft, 420),
            ],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.result_model)
        self.result_table = QTableView()
        configure_table(self.result_table, self.result_model, wrap_headers=True)
        install_copy_menu(self.result_table, self.result_model)
        table_card.layout.addWidget(self.result_table, 1)
        root.addWidget(table_card, 1)

        report_card = Card()
        report_title = QLabel("Report")
        report_title.setObjectName("SectionTitle")
        report_card.layout.addWidget(report_title)
        self.report_output = QPlainTextEdit()
        self.report_output.setReadOnly(True)
        self.report_output.setMaximumHeight(170)
        self.report_output.setPlaceholderText("Summary scan/copy akan muncul di sini.")
        report_card.layout.addWidget(self.report_output)
        root.addWidget(report_card)

        self.register_busy_widgets(
            self.scan_btn,
            self.copy_btn,
            self.export_btn,
            self.open_dest_btn,
            self.clear_btn,
            self.source_picker.button,
            self.destination_picker.button,
            self.part_input,
            self.line_filter_input,
        )
        self._update_actions()

    def set_busy(self, busy, text=None):
        super().set_busy(busy, text)
        if not busy:
            self._update_actions()

    def browse_source_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Source Folder")
        if folder_path:
            self.source_picker.set_path(folder_path)
            self._update_actions()

    def browse_destination_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Destination Folder")
        if folder_path:
            self.destination_picker.set_path(folder_path)
            self._update_actions()

    def scan_programs(self):
        config = CrbProgramCollectorConfig(
            source_folder=self.source_picker.path(),
            destination_folder=self.destination_picker.path(),
            part_numbers_text=self.part_input.toPlainText(),
            line_filter_text=self.line_filter_input.toPlainText(),
        )
        self.scan_result = None
        self.copy_result = None
        self.result_model.set_records([])
        self.report_output.clear()

        worker = None

        def task():
            return scan_crb_programs(
                config,
                progress_callback=lambda percent, message: worker.signals.progress.emit(percent, message),
            )

        worker = TaskWorker(task)
        worker._busy_text = "Scanning .crb files..."
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, "Scanning .crb files..."))
        worker.signals.progress.connect(self._on_progress)
        worker.signals.result.connect(self._on_scan_done)
        worker.signals.error.connect(self._show_worker_error)
        worker.signals.finished.connect(lambda w=worker: self._finish_worker(w))
        self.thread_pool.start(worker)

    def _on_scan_done(self, result):
        self.scan_result = result
        self.copy_result = None
        self.result_model.set_records(self._match_records(result))
        self.report_output.setPlainText(format_crb_program_report(result))

        if result.matches:
            self.status_label.setText(f"Scan complete: {len(result.matches)} file(s) matched")
        elif result.total_crb_scanned == 0:
            self.status_label.setText("No .crb files found")
        else:
            self.status_label.setText("No matching CRB file")
        self._update_actions()

    def copy_matches(self):
        if self.scan_result is None:
            QMessageBox.information(self, "Copy", "Belum ada hasil scan.")
            return

        worker = None

        def task():
            return copy_crb_programs(
                self.scan_result,
                self.destination_picker.path(),
                progress_callback=lambda percent, message: worker.signals.progress.emit(percent, message),
            )

        worker = TaskWorker(task)
        worker._busy_text = "Copying .crb files..."
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, "Copying .crb files..."))
        worker.signals.progress.connect(self._on_progress)
        worker.signals.result.connect(self._on_copy_done)
        worker.signals.error.connect(self._show_worker_error)
        worker.signals.finished.connect(lambda w=worker: self._finish_worker(w))
        self.thread_pool.start(worker)

    def _on_copy_done(self, result):
        self.copy_result = result
        self.report_output.setPlainText(format_crb_program_report(result.scan_result, result))
        self.status_label.setText(f"Copy complete: {result.copied_count}/{len(result.rows)} copied")
        self._update_actions()
        QMessageBox.information(
            self,
            "Copy Complete",
            f"Copied: {result.copied_count}\nError/verify failed: {result.error_count}\nDestination:\n{result.destination_folder}",
        )

    def export_report(self):
        if self.scan_result is None:
            QMessageBox.information(self, "Export Report", "Belum ada report untuk diexport.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save CRB Program Collector Report",
            suggest_report_name(),
            "Text Report (*.txt);;CSV Report (*.csv)",
        )
        if not output_path:
            return

        try:
            saved_path = export_crb_program_report(self.scan_result, output_path, self.copy_result)
        except Exception as exc:
            QMessageBox.warning(self, "Export gagal", str(exc))
            return

        self.status_label.setText(f"Report exported: {Path(saved_path).name}")
        self.status_label.setToolTip(saved_path)
        QMessageBox.information(self, "Export Report", f"Report berhasil dibuat:\n{saved_path}")

    def open_destination_folder(self):
        folder_path = self.destination_picker.path()
        if self.copy_result:
            folder_path = self.copy_result.destination_folder
        if not folder_path or not Path(folder_path).is_dir():
            QMessageBox.information(self, "Open Destination Folder", "Folder tujuan belum ada atau belum dipilih.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(folder_path))

    def clear_result(self):
        self.scan_result = None
        self.copy_result = None
        self.result_model.set_records([])
        self.report_output.clear()
        self.status_label.setText("Ready")
        self.status_label.setToolTip("")
        self._update_actions()

    def _on_progress(self, percent, message):
        self.status_label.setText(message)
        self.progress.setVisible(True)
        self.progress.setRange(0, 100)
        self.progress.setValue(max(0, min(100, percent)))

    def _match_records(self, result):
        records = []
        for index, match in enumerate(result.matches, start=1):
            records.append(
                {
                    "no": index,
                    "file_name": match.file_name,
                    "matched_parts": ", ".join(match.matched_parts),
                    "line": match.line or "-",
                    "size_kb": f"{match.size / 1024:.1f}",
                    "source_folder": match.source_folder,
                    "source_path": match.source_path,
                }
            )
        return records

    def _update_actions(self):
        has_scan = self.scan_result is not None
        has_matches = bool(self.scan_result and self.scan_result.matches)
        self.copy_btn.setEnabled(has_matches and bool(self.destination_picker.path()))
        self.export_btn.setEnabled(has_scan)
        self.open_dest_btn.setEnabled(bool(self.destination_picker.path()) or self.copy_result is not None)
