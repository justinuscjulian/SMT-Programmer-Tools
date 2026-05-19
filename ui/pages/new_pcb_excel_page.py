from pathlib import Path
from datetime import datetime
import tempfile

from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
)

from services.new_pcb_excel_service import NewPcbExcelConfig, generate_new_pcb_excel, suggest_output_name
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker


class NewPcbExcelPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Generate New PCB Program Excel")
        title.setObjectName("SectionTitle")
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        files_card = Card()
        files_title = QLabel("Source Files")
        files_title.setObjectName("SectionTitle")
        files_card.layout.addWidget(files_title)

        self.cad_picker = FilePicker("CAD Data (.txt):")
        self.bom_picker = FilePicker("BOM File (.tsv):")
        self.library_picker = FilePicker("Excel Part Library:")
        self.reference_picker = FilePicker("Excel Referensi:")
        self.gerber_image_picker = FilePicker("Gerber PCB Image:")
        self.gerber_paste_btn = QPushButton("Paste")
        self.cad_picker.browse_requested.connect(self.browse_cad)
        self.bom_picker.browse_requested.connect(self.browse_bom)
        self.library_picker.browse_requested.connect(self.browse_library)
        self.reference_picker.browse_requested.connect(self.browse_reference)
        self.gerber_image_picker.browse_requested.connect(self.browse_gerber_image)
        self.gerber_paste_btn.clicked.connect(self.paste_gerber_image)

        files_card.layout.addWidget(self.cad_picker)
        files_card.layout.addWidget(self.bom_picker)
        files_card.layout.addWidget(self.library_picker)
        files_card.layout.addWidget(self.reference_picker)
        gerber_row = QHBoxLayout()
        gerber_row.setSpacing(8)
        gerber_row.addWidget(self.gerber_image_picker, 1)
        gerber_row.addWidget(self.gerber_paste_btn)
        files_card.layout.addLayout(gerber_row)
        root.addWidget(files_card)

        detail_card = Card()
        detail_title = QLabel("Excel Identity")
        detail_title.setObjectName("SectionTitle")
        detail_card.layout.addWidget(detail_title)

        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(10)
        self.model_input = self._add_field(grid, 0, 0, "Model")
        self.program_part_input = self._add_field(grid, 0, 2, "Part Number Program")
        self.equal_part_input = self._add_field(grid, 1, 0, "Part Number Persamaan")
        self.pcb_part_input = self._add_field(grid, 1, 2, "PCB Part Number")
        self.pcb_revision_input = self._add_field(grid, 2, 0, "PCB Revision Number")
        self.wo_supply_input = self._add_field(grid, 2, 2, "WO Supply")
        self.creator_input = self._add_field(grid, 3, 0, "Nama pembuat")
        self.line_input = self._add_field(grid, 3, 2, "LINE")
        detail_card.layout.addLayout(grid)
        root.addWidget(detail_card)

        action_bar = QHBoxLayout()
        self.generate_btn = QPushButton("Generate Excel")
        self.generate_btn.setObjectName("SuccessButton")
        self.generate_btn.clicked.connect(self.generate_excel)
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
            self.cad_picker.button,
            self.bom_picker.button,
            self.library_picker.button,
            self.reference_picker.button,
            self.gerber_image_picker.button,
            self.gerber_paste_btn,
        )

    def _add_field(self, grid, row, column, label_text):
        label = QLabel(label_text)
        label.setObjectName("MutedLabel")
        line_edit = QLineEdit()
        grid.addWidget(label, row, column)
        grid.addWidget(line_edit, row, column + 1)
        return line_edit

    def browse_cad(self):
        self._browse_file(self.cad_picker, "Select CAD Data", "CAD Text (*.txt);;All Files (*)")

    def browse_bom(self):
        self._browse_file(self.bom_picker, "Select BOM File", "BOM (*.tsv *.xlsx *.xls);;All Files (*)")

    def browse_library(self):
        self._browse_file(self.library_picker, "Select Excel Part Library", "Excel (*.xlsb *.xlsx *.xls);;All Files (*)")

    def browse_reference(self):
        self._browse_file(self.reference_picker, "Select Excel Referensi", "Excel (*.xlsx *.xlsm *.xls *.xlsb);;All Files (*)")

    def browse_gerber_image(self):
        self._browse_file(self.gerber_image_picker, "Select Gerber PCB Image", "Images (*.png *.jpg *.jpeg *.bmp *.gif);;All Files (*)")

    def paste_gerber_image(self):
        clipboard = QApplication.clipboard()
        image = clipboard.image()
        if image.isNull():
            pixmap = clipboard.pixmap()
            if not pixmap.isNull():
                image = pixmap.toImage()

        if image.isNull():
            QMessageBox.warning(self, "Clipboard kosong", "Tidak ada gambar di clipboard.")
            return

        filename = f"smt_tools_gerber_clipboard_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.png"
        image_path = Path(tempfile.gettempdir()) / filename
        if not image.save(str(image_path), "PNG"):
            QMessageBox.warning(self, "Paste gagal", "Gambar clipboard tidak bisa disimpan sementara.")
            return

        self.gerber_image_picker.set_path(str(image_path), "Clipboard image pasted")
        self.status_label.setText("Gerber image pasted from clipboard")

    def _browse_file(self, picker, title, file_filter):
        file_path, _ = QFileDialog.getOpenFileName(self, title, "", file_filter)
        if file_path:
            picker.set_path(file_path)

    def generate_excel(self):
        config = self._config(output_path="")
        suggested_name = suggest_output_name(config)
        output_path, _ = QFileDialog.getSaveFileName(self, "Save Generated Excel", suggested_name, "Excel Workbook (*.xlsx)")
        if not output_path:
            return
        config = self._config(output_path=output_path)
        self.run_worker(lambda: generate_new_pcb_excel(config), self._on_generate_done, "Generating NEW PCB Excel...")

    def _on_generate_done(self, output_path):
        self.status_label.setText(f"Saved: {Path(output_path).name}")
        self.status_label.setToolTip(output_path)
        QMessageBox.information(self, "Generate Excel", "Silahkan lanjutkan proses pada Excel")

    def _config(self, output_path):
        return NewPcbExcelConfig(
            cad_path=self.cad_picker.path(),
            bom_path=self.bom_picker.path(),
            library_path=self.library_picker.path(),
            reference_path=self.reference_picker.path(),
            gerber_image_path=self.gerber_image_picker.path(),
            model=self.model_input.text(),
            program_part_number=self.program_part_input.text(),
            equivalent_part_number=self.equal_part_input.text(),
            pcb_part_number=self.pcb_part_input.text(),
            pcb_revision=self.pcb_revision_input.text(),
            wo_supply=self.wo_supply_input.text(),
            creator=self.creator_input.text(),
            line=self.line_input.text(),
            output_path=output_path,
        )

    def clear_form(self):
        for picker in (self.cad_picker, self.bom_picker, self.library_picker, self.reference_picker, self.gerber_image_picker):
            picker.clear()
        for field in (
            self.model_input,
            self.program_part_input,
            self.equal_part_input,
            self.pcb_part_input,
            self.pcb_revision_input,
            self.wo_supply_input,
            self.creator_input,
            self.line_input,
        ):
            field.clear()
        self.status_label.setText("")
        self.status_label.setToolTip("")
