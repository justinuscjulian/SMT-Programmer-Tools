from pathlib import Path

from PySide6.QtWidgets import (
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
)

from services.insert_point_service import InsertPointConfig, generate_insert_point, suggest_output_name
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker


class InsertPointPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Get Insert Point")
        title.setObjectName("SectionTitle")
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        source_card = Card()
        source_title = QLabel("Source Data")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)

        self.plan_picker = FilePicker("Excel PLAN:")
        self.pcb_folder_picker = FilePicker("Folder Induk PCB:")
        self.plan_picker.browse_requested.connect(self.browse_plan)
        self.pcb_folder_picker.browse_requested.connect(self.browse_pcb_folder)
        source_card.layout.addWidget(self.plan_picker)
        source_card.layout.addWidget(self.pcb_folder_picker)
        root.addWidget(source_card)

        range_card = Card()
        range_title = QLabel("Row Range")
        range_title.setObjectName("SectionTitle")
        range_card.layout.addWidget(range_title)

        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(10)
        self.start_row_spin = self._add_spinbox(grid, 0, 0, "Start Row", 2)
        self.end_row_spin = self._add_spinbox(grid, 0, 2, "End Row", 100)
        range_card.layout.addLayout(grid)
        root.addWidget(range_card)

        action_bar = QHBoxLayout()
        self.generate_btn = QPushButton("Generate Insert Point")
        self.generate_btn.setObjectName("SuccessButton")
        self.generate_btn.clicked.connect(self.generate_insert_point)
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_form)
        action_bar.addWidget(self.generate_btn)
        action_bar.addWidget(self.clear_btn)
        action_bar.addStretch(1)
        root.addLayout(action_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)
        root.addStretch(1)

        self.register_busy_widgets(
            self.generate_btn,
            self.clear_btn,
            self.plan_picker.button,
            self.pcb_folder_picker.button,
        )

    def _add_spinbox(self, grid, row, column, label_text, value):
        label = QLabel(label_text)
        label.setObjectName("MutedLabel")
        spinbox = QSpinBox()
        spinbox.setRange(1, 1048576)
        spinbox.setValue(value)
        grid.addWidget(label, row, column)
        grid.addWidget(spinbox, row, column + 1)
        return spinbox

    def browse_plan(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Pilih file Excel PLAN",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls *.xlsb);;All Files (*)",
        )
        if file_path:
            self.plan_picker.set_path(file_path)

    def browse_pcb_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Pilih Folder Induk PCB")
        if folder_path:
            self.pcb_folder_picker.set_path(folder_path)

    def generate_insert_point(self):
        if not self._validate_before_save():
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Insert Point Data",
            suggest_output_name(),
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        config = self._config(output_path)
        self.run_worker(
            lambda: generate_insert_point(config),
            self._on_generate_done,
            "Generating Insert Point data...",
        )

    def _on_generate_done(self, result):
        output_name = Path(result.output_path).name
        self.status_label.setText(f"Saved: {output_name}")
        self.status_label.setToolTip(result.output_path)
        QMessageBox.information(
            self,
            "Insert Point Selesai",
            f"Berhasil: {result.success_count}\nError: {result.error_count}\nFile: {result.output_path}",
        )

    def _config(self, output_path):
        return InsertPointConfig(
            plan_path=self.plan_picker.path(),
            main_folder=self.pcb_folder_picker.path(),
            start_row=self.start_row_spin.value(),
            end_row=self.end_row_spin.value(),
            output_path=output_path,
        )

    def _validate_before_save(self):
        plan_path = self.plan_picker.path()
        folder_path = self.pcb_folder_picker.path()

        if not plan_path:
            QMessageBox.warning(self, "Input belum lengkap", "File Excel PLAN belum dipilih.")
            return False
        if not Path(plan_path).is_file():
            QMessageBox.warning(self, "File tidak ditemukan", f"File Excel PLAN tidak ditemukan:\n{plan_path}")
            return False
        if not folder_path:
            QMessageBox.warning(self, "Input belum lengkap", "Folder Induk PCB belum dipilih.")
            return False
        if not Path(folder_path).is_dir():
            QMessageBox.warning(self, "Folder tidak ditemukan", f"Folder Induk PCB tidak ditemukan:\n{folder_path}")
            return False
        if self.end_row_spin.value() < self.start_row_spin.value():
            QMessageBox.warning(self, "Range row tidak valid", "End Row tidak boleh lebih kecil dari Start Row.")
            return False

        return True

    def clear_form(self):
        self.plan_picker.clear()
        self.pcb_folder_picker.clear()
        self.start_row_spin.setValue(2)
        self.end_row_spin.setValue(100)
        self.status_label.setText("")
        self.status_label.setToolTip("")
