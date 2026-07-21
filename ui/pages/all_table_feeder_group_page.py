from pathlib import Path
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QTableView,
    QPlainTextEdit,
    QLineEdit,
    QComboBox,
    QProgressBar,
    QSpinBox
)

from models.table_model import ColumnSpec, RecordTableModel
from services.all_table_feeder_group_service import generate_all_table_groups, export_all_table_groups
from ui.pages.base import WorkerPage, TaskWorker
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu


class AllTableFeederGroupPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.groups_result = None
        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("All Table Fix Feeder Generator")
        title.setObjectName("SectionTitle")
        self.summary_label = QLabel("0 GROUPS")
        self.summary_label.setObjectName("MutedLabel")
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setFixedWidth(200)
        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addWidget(self.summary_label)
        header.addStretch(1)
        header.addWidget(self.progress)
        header.addWidget(self.status_label)
        root.addLayout(header)

        source_card = Card()
        source_title = QLabel("Sources & Configuration")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)

        self.folder_picker = FilePicker("Folder Induk Program Excel:")
        self.folder_picker.browse_requested.connect(self.browse_source_folder)
        source_card.layout.addWidget(self.folder_picker)

        self.master_picker = FilePicker("Master Mapping Excel (Multiple Feeder Output):")
        self.master_picker.browse_requested.connect(self.browse_master_file)
        source_card.layout.addWidget(self.master_picker)

        line_type_layout = QHBoxLayout()
        line_type_label = QLabel("Pilih Tipe Line:")
        self.line_type_combo = QComboBox()
        self.line_type_combo.addItems(["Line 1-5", "Line 6-7", "Line 8"])
        line_type_layout.addWidget(line_type_label)
        line_type_layout.addWidget(self.line_type_combo)
        line_type_layout.addStretch()
        source_card.layout.addLayout(line_type_layout)

        list_label = QLabel("List PCB Part Number (Opsional, kosongkan untuk proses semua di folder):")
        list_label.setObjectName("MutedLabel")
        source_card.layout.addWidget(list_label)
        
        self.list_input = QPlainTextEdit()
        self.list_input.setPlaceholderText("Paste list PCB Part Number di sini, pisahkan dengan enter atau koma...")
        self.list_input.setMinimumHeight(100)
        source_card.layout.addWidget(self.list_input)

        option_row = QHBoxLayout()
        similarity_label = QLabel("Min Similarity (%):")
        similarity_label.setMinimumWidth(150)
        self.similarity_spin = QSpinBox()
        self.similarity_spin.setRange(1, 100)
        self.similarity_spin.setValue(70)
        self.similarity_spin.setSuffix("%")
        option_row.addWidget(similarity_label)
        option_row.addWidget(self.similarity_spin)
        option_row.addStretch(1)
        source_card.layout.addLayout(option_row)

        option_row2 = QHBoxLayout()
        shared_label = QLabel("Min Shared Components:")
        shared_label.setMinimumWidth(150)
        self.shared_spin = QSpinBox()
        self.shared_spin.setRange(1, 1000)
        self.shared_spin.setValue(20)
        option_row2.addWidget(shared_label)
        option_row2.addWidget(self.shared_spin)
        option_row2.addStretch(1)
        source_card.layout.addLayout(option_row2)

        root.addWidget(source_card)

        action_bar = QHBoxLayout()
        self.analyze_btn = QPushButton("Generate Groups")
        self.analyze_btn.setObjectName("PrimaryButton")
        self.analyze_btn.clicked.connect(self.analyze_groups)
        
        self.export_btn = QPushButton("Export Excel")
        self.export_btn.setObjectName("SuccessButton")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_results)
        
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_results)
        
        action_bar.addWidget(self.analyze_btn)
        action_bar.addWidget(self.export_btn)
        action_bar.addStretch(1)
        action_bar.addWidget(self.clear_btn)
        root.addLayout(action_bar)

        self.register_busy_widgets(
            self.analyze_btn,
            self.export_btn,
            self.clear_btn,
            self.folder_picker.button,
            self.master_picker.button,
            self.line_type_combo,
            self.list_input,
        )

    def browse_source_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Pilih Folder Induk PCB")
        if folder_path:
            self.folder_picker.set_path(folder_path)

    def browse_master_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Pilih File Master Mapping", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.master_picker.set_path(file_path)

    def analyze_groups(self):
        if not self._validate_before_analysis():
            return

        crb_folder = self.folder_picker.path()
        master_excel = self.master_picker.path()
        target_list_text = self.list_input.toPlainText()
        line_type = self.line_type_combo.currentText()
        min_sim = self.similarity_spin.value()
        min_shared = self.shared_spin.value()

        self.groups_result = None
        self.export_btn.setEnabled(False)

        worker = None

        def task():
            return generate_all_table_groups(
                crb_folder, master_excel, target_list_text, line_type, min_sim, min_shared, worker.signals.progress.emit
            )

        worker = TaskWorker(task)
        worker._busy_text = "Generating Fix Feeder Groups..."
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, "Generating Fix Feeder Groups..."))
        worker.signals.progress.connect(self._update_progress)
        worker.signals.result.connect(self._on_analysis_done)
        worker.signals.error.connect(self._show_worker_error)
        worker.signals.finished.connect(lambda w=worker: self._finish_worker(w))
        self.thread_pool.start(worker)

    def _on_analysis_done(self, result):
        self.groups_result = result
        self.export_btn.setEnabled(True)

        if not result:
            self.status_label.setText("Tidak ada grup yang dihasilkan")
            QMessageBox.information(self, "Data kosong", "Tidak ada PCB yang valid di folder tersebut.")
        else:
            self.status_label.setText(f"Done: {len(result)} groups generated.")
            self.summary_label.setText(f"{len(result)} GROUPS")
            QMessageBox.information(self, "Selesai", f"Berhasil membuat {len(result)} Fix Feeder groups. Silahkan Export ke Excel.")

    def export_results(self):
        if not self.groups_result:
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save All Table Fix Feeder Groups",
            "All_Table_Fix_Feeder_Groups.xlsx",
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        try:
            saved_path = export_all_table_groups(self.groups_result, output_path)
        except Exception as exc:
            QMessageBox.warning(self, "Export gagal", str(exc))
            return

        self.status_label.setText(f"Exported: {Path(saved_path).name}")
        self.status_label.setToolTip(saved_path)
        QMessageBox.information(self, "Export Excel", f"File berhasil dibuat:\n{saved_path}")

    def clear_results(self):
        self.groups_result = None
        self.folder_picker.clear()
        self.master_picker.clear()
        self.list_input.clear()
        self.export_btn.setEnabled(False)
        self.summary_label.setText("0 GROUPS")
        self.status_label.setText("Ready")
        self.status_label.setToolTip("")
        self.progress.setVisible(False)
        self.progress.setValue(0)

    def _update_progress(self, percent, text):
        self.progress.setValue(percent)
        self.status_label.setText(text)

    def _validate_before_analysis(self):
        folder_path = self.folder_picker.path()
        if not folder_path or not Path(folder_path).is_dir():
            QMessageBox.warning(self, "Input belum lengkap", "Folder Induk .crb / .txt belum dipilih atau tidak valid.")
            return False
            
        master_path = self.master_picker.path()
        if not master_path or not Path(master_path).is_file():
            QMessageBox.warning(self, "Input belum lengkap", "File Master Mapping Excel belum dipilih atau tidak valid.")
            return False
            
        return True
