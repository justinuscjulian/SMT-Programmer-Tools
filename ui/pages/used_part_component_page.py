from pathlib import Path

from PySide6.QtWidgets import (
    QButtonGroup,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
)

from services.used_part_component_service import (
    MODE_PCB_LIST,
    MODE_PROGRAM_FOLDER,
    UsedPartComponentConfig,
    generate_used_part_component,
    suggest_output_name,
)
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker


class UsedPartComponentPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Used Part Component")
        title.setObjectName("SectionTitle")
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        mode_card = Card()
        mode_title = QLabel("Collect Mode")
        mode_title.setObjectName("SectionTitle")
        mode_card.layout.addWidget(mode_title)

        mode_row = QHBoxLayout()
        self.mode_1_radio = QRadioButton("Mode 1 - PCB Part Number List")
        self.mode_2_radio = QRadioButton("Mode 2 - One PCB Program Folder")
        self.mode_2_radio.setChecked(True)
        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.mode_1_radio)
        self.mode_group.addButton(self.mode_2_radio)
        self.mode_1_radio.toggled.connect(self._sync_mode_ui)
        mode_row.addWidget(self.mode_1_radio)
        mode_row.addWidget(self.mode_2_radio)
        mode_row.addStretch(1)
        mode_card.layout.addLayout(mode_row)
        root.addWidget(mode_card)

        source_card = Card()
        source_title = QLabel("Source Data")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)

        self.source_picker = FilePicker("Folder Program:")
        self.source_picker.browse_requested.connect(self.browse_source_folder)
        source_card.layout.addWidget(self.source_picker)

        self.pcb_list_label = QLabel("PCB Part Number List")
        self.pcb_list_label.setObjectName("MutedLabel")
        self.pcb_list_input = QPlainTextEdit()
        self.pcb_list_input.setPlaceholderText("Paste PCB part number, satu baris satu part number")
        self.pcb_list_input.setMinimumHeight(110)
        source_card.layout.addWidget(self.pcb_list_label)
        source_card.layout.addWidget(self.pcb_list_input)
        root.addWidget(source_card)

        action_bar = QHBoxLayout()
        self.generate_btn = QPushButton("Generate Used Part Excel")
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
            self.source_picker.button,
            self.pcb_list_input,
            self.mode_1_radio,
            self.mode_2_radio,
        )

        self._sync_mode_ui()

    def browse_source_folder(self):
        title = "Pilih Folder Induk PCB" if self._selected_mode() == MODE_PCB_LIST else "Pilih Folder Program PCB"
        folder_path = QFileDialog.getExistingDirectory(self, title)
        if folder_path:
            self.source_picker.set_path(folder_path)

    def generate_excel(self):
        if not self._validate_before_save():
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Used Part Component",
            suggest_output_name(self._selected_mode()),
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        config = self._config(output_path)
        self.run_worker(
            lambda: generate_used_part_component(config),
            self._on_generate_done,
            "Generating Used Part Component...",
        )

    def _on_generate_done(self, result):
        output_name = Path(result.output_path).name
        self.status_label.setText(f"Saved: {output_name}")
        self.status_label.setToolTip(result.output_path)

        skipped_text = f"\nSkipped: {len(result.skipped_files)}" if result.skipped_files else ""
        QMessageBox.information(
            self,
            "Used Part Component Selesai",
            f"Group: {result.group_count}\nFile dibaca: {result.file_count}\nUnique part: {result.part_count}{skipped_text}\nFile: {result.output_path}",
        )

    def _validate_before_save(self):
        source_folder = self.source_picker.path()
        if not source_folder:
            QMessageBox.warning(self, "Input belum lengkap", "Folder source belum dipilih.")
            return False
        if not Path(source_folder).is_dir():
            QMessageBox.warning(self, "Folder tidak ditemukan", f"Folder source tidak ditemukan:\n{source_folder}")
            return False
        if self._selected_mode() == MODE_PCB_LIST and not self.pcb_list_input.toPlainText().strip():
            QMessageBox.warning(self, "Input belum lengkap", "List PCB Part Number belum diisi.")
            return False
        return True

    def _config(self, output_path):
        return UsedPartComponentConfig(
            mode=self._selected_mode(),
            source_folder=self.source_picker.path(),
            pcb_part_numbers=self.pcb_list_input.toPlainText(),
            output_path=output_path,
        )

    def _selected_mode(self):
        return MODE_PCB_LIST if self.mode_1_radio.isChecked() else MODE_PROGRAM_FOLDER

    def _sync_mode_ui(self):
        is_mode_1 = self._selected_mode() == MODE_PCB_LIST
        self.source_picker.label.setText("Folder Induk PCB:" if is_mode_1 else "Folder Program:")
        self.pcb_list_label.setVisible(is_mode_1)
        self.pcb_list_input.setVisible(is_mode_1)
        self.source_picker.clear()
        self.status_label.setText("")
        self.status_label.setToolTip("")

    def clear_form(self):
        self.source_picker.clear()
        self.pcb_list_input.clear()
        self.status_label.setText("")
        self.status_label.setToolTip("")
