from pathlib import Path

from PySide6.QtWidgets import (
    QButtonGroup,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
)

from services.plan_service import (
    PLAN_TYPE_FIRST,
    PLAN_TYPE_SECOND,
    PLAN_TYPE_THIRD,
    PlanConfig,
    generate_plan,
    suggest_output_name,
)
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker


class PlanPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedLabel")
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        mode_card = Card()
        mode_title = QLabel("Plan Type")
        mode_title.setObjectName("SectionTitle")
        mode_card.layout.addWidget(mode_title)

        mode_row = QHBoxLayout()
        self.first_plan_btn = QPushButton("1ST PLAN")
        self.first_plan_btn.setObjectName("SegmentedButton")
        self.first_plan_btn.setCheckable(True)
        self.second_plan_btn = QPushButton("2ND PLAN")
        self.second_plan_btn.setObjectName("SegmentedButton")
        self.second_plan_btn.setCheckable(True)
        self.third_plan_btn = QPushButton("3RD PLAN")
        self.third_plan_btn.setObjectName("SegmentedButton")
        self.third_plan_btn.setCheckable(True)

        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)
        self.mode_group.addButton(self.first_plan_btn)
        self.mode_group.addButton(self.second_plan_btn)
        self.mode_group.addButton(self.third_plan_btn)
        self.first_plan_btn.setChecked(True)

        mode_row.addWidget(self.first_plan_btn)
        mode_row.addWidget(self.second_plan_btn)
        mode_row.addWidget(self.third_plan_btn)
        mode_row.addStretch(1)
        mode_card.layout.addLayout(mode_row)
        root.addWidget(mode_card)

        files_card = Card()
        files_title = QLabel("Source Files")
        files_title.setObjectName("SectionTitle")
        files_card.layout.addWidget(files_title)

        self.previous_plan_picker = FilePicker("PLAN sebelumnya:")
        self.new_plan_picker = FilePicker("PLAN baru:")
        self.history_folder_picker = FilePicker("Folder history:")
        self.previous_plan_picker.browse_requested.connect(self.browse_previous_plan)
        self.new_plan_picker.browse_requested.connect(self.browse_new_plan)
        self.history_folder_picker.browse_requested.connect(self.browse_history_folder)
        files_card.layout.addWidget(self.previous_plan_picker)
        files_card.layout.addWidget(self.new_plan_picker)
        files_card.layout.addWidget(self.history_folder_picker)
        root.addWidget(files_card)

        action_bar = QHBoxLayout()
        self.generate_btn = QPushButton("Generate PLAN")
        self.generate_btn.setObjectName("SuccessButton")
        self.generate_btn.clicked.connect(self.generate)
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
            self.first_plan_btn,
            self.second_plan_btn,
            self.third_plan_btn,
            self.generate_btn,
            self.clear_btn,
            self.previous_plan_picker.button,
            self.new_plan_picker.button,
            self.history_folder_picker.button,
        )

    def browse_previous_plan(self):
        self._browse_file(self.previous_plan_picker, "Pilih PLAN sebelumnya")

    def browse_new_plan(self):
        self._browse_file(self.new_plan_picker, "Pilih PLAN baru")

    def browse_history_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Pilih Folder History Program")
        if folder_path:
            self.history_folder_picker.set_path(folder_path)

    def _browse_file(self, picker, title):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            title,
            "",
            "Excel Files (*.xlsx *.xlsm *.xls *.xlsb);;All Files (*)",
        )
        if file_path:
            picker.set_path(file_path)

    def generate(self):
        if not self._validate_before_save():
            return

        plan_type = self._plan_type()
        suggested_name = suggest_output_name(
            plan_type,
            self.previous_plan_picker.path(),
            self.new_plan_picker.path(),
        )
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Generated PLAN",
            suggested_name,
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        config = PlanConfig(
            plan_type=plan_type,
            previous_plan_path=self.previous_plan_picker.path(),
            new_plan_path=self.new_plan_picker.path(),
            output_path=output_path,
            history_folder_path=self.history_folder_picker.path(),
        )
        self.run_worker(lambda: generate_plan(config), self._on_generate_done, "Generating PLAN...")

    def _on_generate_done(self, result):
        output_name = Path(result.output_path).name
        self.status_label.setText(f"Saved: {output_name}")
        self.status_label.setToolTip(result.output_path)
        QMessageBox.information(
            self,
            "PLAN Selesai",
            f"Sheet: {result.sheet_name}\n"
            f"Match PLAN sebelumnya: {result.matched_count}\n"
            f"Perlu history: {result.not_found_count}\n"
            f"Terisi dari history: {result.history_found_count}\n"
            f"Ada di line lain: {result.history_other_line_count}\n"
            f"History tidak ditemukan: {result.history_not_found_count}\n"
            f"File: {result.output_path}",
        )

    def _plan_type(self):
        if self.second_plan_btn.isChecked():
            return PLAN_TYPE_SECOND
        if self.third_plan_btn.isChecked():
            return PLAN_TYPE_THIRD
        return PLAN_TYPE_FIRST

    def _validate_before_save(self):
        for picker, label in (
            (self.previous_plan_picker, "PLAN sebelumnya"),
            (self.new_plan_picker, "PLAN baru"),
        ):
            path = picker.path()
            if not path:
                QMessageBox.warning(self, "Input belum lengkap", f"{label} belum dipilih.")
                return False
            if not Path(path).is_file():
                QMessageBox.warning(self, "File tidak ditemukan", f"{label} tidak ditemukan:\n{path}")
                return False

        history_path = self.history_folder_picker.path()
        if not history_path:
            QMessageBox.warning(self, "Input belum lengkap", "Folder history belum dipilih.")
            return False
        if not Path(history_path).is_dir():
            QMessageBox.warning(self, "Folder tidak ditemukan", f"Folder history tidak ditemukan:\n{history_path}")
            return False

        return True

    def clear_form(self):
        self.previous_plan_picker.clear()
        self.new_plan_picker.clear()
        self.history_folder_picker.clear()
        self.first_plan_btn.setChecked(True)
        self.status_label.setText("")
        self.status_label.setToolTip("")
