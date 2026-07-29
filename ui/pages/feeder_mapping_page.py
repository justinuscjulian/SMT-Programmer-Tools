import re
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QTableView,
    QVBoxLayout,
)

from models.table_model import ColumnSpec, RecordTableModel
from services import feeder_mapping_service, npm_txt_editor_service
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu


class NpmTxtSlotEditDialog(QDialog):
    def __init__(self, parent=None, table=1, slot=1, pos="", part_number="", feeder_id=""):
        super().__init__(parent)
        self.setWindowTitle("Edit / Tambah Slot Feeder NPM")
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        self.table_combo = QComboBox()
        for t in range(1, 11):
            self.table_combo.addItem(f"Table {t}", t)
        self.table_combo.setCurrentIndex(max(0, table - 1))

        self.slot_spin = QSpinBox()
        self.slot_spin.setRange(1, 30)
        self.slot_spin.setValue(slot)

        self.pos_combo = QComboBox()
        self.pos_combo.addItem("L (Left Position)", "L")
        self.pos_combo.addItem("R (Right Position)", "R")
        self.pos_combo.addItem("Single / Full Slot (No L/R)", "")
        p_upper = pos.strip().upper()
        if p_upper == "L":
            self.pos_combo.setCurrentIndex(0)
        elif p_upper == "R":
            self.pos_combo.setCurrentIndex(1)
        else:
            self.pos_combo.setCurrentIndex(2)

        self.part_input = QLineEdit(part_number)
        self.part_input.setPlaceholderText("Contoh: EAE32246001 / 0RJ1000")

        self.feeder_input = QLineEdit(feeder_id)
        self.feeder_input.setPlaceholderText("Feeder ID (opsional, kosongkan untuk auto-infer)")

        form_layout.addRow("Table:", self.table_combo)
        form_layout.addRow("Slot Number:", self.slot_spin)
        form_layout.addRow("Position (L/R/Single):", self.pos_combo)
        form_layout.addRow("Part Number:", self.part_input)
        form_layout.addRow("Feeder ID:", self.feeder_input)

        layout.addLayout(form_layout)

        btn_box = QHBoxLayout()
        ok_btn = QPushButton("Simpan Slot")
        ok_btn.setObjectName("PrimaryButton")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Batal")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addStretch(1)
        btn_box.addWidget(ok_btn)
        btn_box.addWidget(cancel_btn)

        layout.addLayout(btn_box)

    def get_data(self):
        return {
            "table": self.table_combo.currentData(),
            "slot": self.slot_spin.value(),
            "pos": self.pos_combo.currentData(),
            "part_number": self.part_input.text().strip().upper(),
            "feeder_id": self.feeder_input.text().strip(),
        }


class FeederMappingPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.mapping_result = None
        self.source_path = ""
        self.multi_source_paths = []
        self.npm_doc = None
        self.npm_doc_slots = []
        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Generate Feeder Mapping")
        title.setObjectName("SectionTitle")
        self.summary_label = QLabel("0 ROWS")
        self.summary_label.setObjectName("MutedLabel")
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addWidget(self.summary_label)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        mode_card = Card()
        mode_title = QLabel("Generator Mode")
        mode_title.setObjectName("SectionTitle")
        mode_card.layout.addWidget(mode_title)

        mode_grid = QGridLayout()
        mode_grid.setContentsMargins(4, 4, 4, 4)
        mode_grid.setHorizontalSpacing(28)
        mode_grid.setVerticalSpacing(12)

        self.single_mode_radio = QRadioButton("Single Feeder File")
        self.multiple_mode_radio = QRadioButton("Multiple Feeder Files")
        self.npm_editor_mode_radio = QRadioButton("Edit NPM Feeder TXT")

        self.import_mode_radio = QRadioButton("Excel to NPM Feeder TXT")
        self.import_mode_line8_radio = QRadioButton("Excel to NPM Feeder TXT (Line 8)")
        self.import_mode_khusus_sub_radio = QRadioButton("Excel to NPM Feeder TXT (Khusus Sub)")
        self.group_import_mode_radio = QRadioButton("Fix Feeder Group to NPM TXT")
        self.group_import_mode_line8_radio = QRadioButton("Fix Feeder Group to NPM TXT (Line 8)")
        self.group_import_mode_khusus_sub_radio = QRadioButton("Fix Feeder Group to NPM TXT (Khusus Sub)")

        self.cm602_mode_radio = QRadioButton("CM602")
        self.cm602_program_cm_txt_mode_radio = QRadioButton("CM602 Program File Converter to CM.txt")
        self.cm602_feeder_fix_mode_radio = QRadioButton("Excel to CM602 FeederFix TXT")

        self.single_mode_radio.setChecked(True)
        self.mode_group = QButtonGroup(self)

        radios = [
            self.single_mode_radio,
            self.multiple_mode_radio,
            self.npm_editor_mode_radio,
            self.import_mode_radio,
            self.import_mode_line8_radio,
            self.import_mode_khusus_sub_radio,
            self.group_import_mode_radio,
            self.group_import_mode_line8_radio,
            self.group_import_mode_khusus_sub_radio,
            self.cm602_mode_radio,
            self.cm602_program_cm_txt_mode_radio,
            self.cm602_feeder_fix_mode_radio,
        ]
        for rb in radios:
            self.mode_group.addButton(rb)
            rb.toggled.connect(self._sync_mode_ui)

        # Row 0: Standard NPM Modes
        mode_grid.addWidget(self.single_mode_radio, 0, 0)
        mode_grid.addWidget(self.multiple_mode_radio, 0, 1)
        mode_grid.addWidget(self.npm_editor_mode_radio, 0, 2)

        # Row 1: NPM Excel Import Modes (Standard & Line 8)
        mode_grid.addWidget(self.import_mode_radio, 1, 0)
        mode_grid.addWidget(self.import_mode_line8_radio, 1, 1)
        mode_grid.addWidget(self.group_import_mode_radio, 1, 2)
        mode_grid.addWidget(self.group_import_mode_line8_radio, 1, 3)

        # Row 2: Khusus Sub Modes
        mode_grid.addWidget(self.import_mode_khusus_sub_radio, 2, 0)
        mode_grid.addWidget(self.group_import_mode_khusus_sub_radio, 2, 1)

        # Row 3: CM602 Modes
        mode_grid.addWidget(self.cm602_mode_radio, 3, 0)
        mode_grid.addWidget(self.cm602_program_cm_txt_mode_radio, 3, 1)
        mode_grid.addWidget(self.cm602_feeder_fix_mode_radio, 3, 2)

        mode_card.layout.addLayout(mode_grid)
        root.addWidget(mode_card)

        source_card = Card()
        source_title = QLabel("Source File(s)")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)
        self.source_picker = FilePicker("NPM Machine Export (.txt/.crb):")
        self.source_picker.browse_requested.connect(self.browse_source)
        source_card.layout.addWidget(self.source_picker)
        self.reference_picker = FilePicker("Reference Feeder Folder:")
        self.reference_picker.browse_requested.connect(self.browse_reference_folder)
        source_card.layout.addWidget(self.reference_picker)
        self.balancing_label = QLabel("Extra Balancing Part Numbers (Optional)")
        self.balancing_label.setObjectName("MutedLabel")
        self.balancing_input = QPlainTextEdit()
        self.balancing_input.setPlainText(feeder_mapping_service.default_balancing_part_numbers_text())
        self.balancing_input.setPlaceholderText("Optional: tambah part number yang mau dipaksa tetap multi-feeder. Auto-detect tetap jalan dari reference folder.")
        self.balancing_input.setMinimumHeight(92)
        source_card.layout.addWidget(self.balancing_label)
        source_card.layout.addWidget(self.balancing_input)
        root.addWidget(source_card)

        action_bar = QHBoxLayout()
        self.preview_btn = QPushButton("Preview Mapping")
        self.preview_btn.setObjectName("PrimaryButton")
        self.preview_btn.clicked.connect(self.preview_mapping)
        self.generate_btn = QPushButton("Generate Excel")
        self.generate_btn.setObjectName("SuccessButton")
        self.generate_btn.setEnabled(False)
        self.generate_btn.clicked.connect(self.generate_excel)
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_data)

        self.editor_add_btn = QPushButton("➕ Tambah Slot Feeder Baru")
        self.editor_add_btn.setObjectName("PrimaryButton")
        self.editor_add_btn.clicked.connect(self.editor_add_slot)
        self.editor_edit_btn = QPushButton("✏️ Edit Slot Terpilih")
        self.editor_edit_btn.clicked.connect(self.editor_edit_slot)
        self.editor_delete_btn = QPushButton("🗑️ Hapus Slot Terpilih")
        self.editor_delete_btn.setObjectName("DangerButton")
        self.editor_delete_btn.clicked.connect(self.editor_delete_slot)
        self.editor_save_btn = QPushButton("💾 Simpan / Export File NPM TXT")
        self.editor_save_btn.setObjectName("SuccessButton")
        self.editor_save_btn.clicked.connect(self.editor_save_file)

        action_bar.addWidget(self.preview_btn)
        action_bar.addWidget(self.generate_btn)
        action_bar.addWidget(self.editor_add_btn)
        action_bar.addWidget(self.editor_edit_btn)
        action_bar.addWidget(self.editor_delete_btn)
        action_bar.addWidget(self.editor_save_btn)
        action_bar.addStretch(1)
        action_bar.addWidget(self.clear_btn)
        root.addLayout(action_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        table_card = Card()
        table_header = QHBoxLayout()
        table_title = QLabel("Detailed Feeder Setup")
        table_title.setObjectName("SectionTitle")
        search_label = QLabel("Search Preview:")
        search_label.setObjectName("MutedLabel")
        self.preview_search_input = QLineEdit()
        self.preview_search_input.setPlaceholderText("Search table, slot, position, location code, atau part number")
        self.preview_search_input.setClearButtonEnabled(True)
        self.preview_search_input.textChanged.connect(self.apply_preview_search)
        self.preview_search_input.setMinimumWidth(360)
        table_header.addWidget(table_title)
        table_header.addStretch(1)
        table_header.addWidget(search_label)
        table_header.addWidget(self.preview_search_input, 1)
        table_card.layout.addLayout(table_header)
        self.mapping_model = RecordTableModel(
            [
                ColumnSpec("table", "Table", Qt.AlignCenter, 90),
                ColumnSpec("slot", "Slot", Qt.AlignCenter, 80),
                ColumnSpec("position", "Position", Qt.AlignCenter, 170),
                ColumnSpec("location_code", "Location Code", Qt.AlignLeft, 170),
                ColumnSpec("part_number", "Part Number", Qt.AlignLeft, 220),
            ],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.mapping_model)
        self.mapping_table = QTableView()
        configure_table(self.mapping_table, self.mapping_model, wrap_headers=True)
        install_copy_menu(self.mapping_table, self.mapping_model)
        table_card.layout.addWidget(self.mapping_table, 1)
        root.addWidget(table_card, 1)

        self.register_busy_widgets(
            self.preview_btn,
            self.generate_btn,
            self.clear_btn,
            self.editor_add_btn,
            self.editor_edit_btn,
            self.editor_delete_btn,
            self.editor_save_btn,
            self.preview_search_input,
            self.source_picker.button,
            self.reference_picker.button,
            self.balancing_input,
            self.single_mode_radio,
            self.multiple_mode_radio,
            self.cm602_mode_radio,
            self.cm602_program_cm_txt_mode_radio,
            self.cm602_feeder_fix_mode_radio,
            self.import_mode_radio,
            self.import_mode_line8_radio,
            self.import_mode_khusus_sub_radio,
            self.group_import_mode_radio,
            self.group_import_mode_line8_radio,
            self.group_import_mode_khusus_sub_radio,
            self.npm_editor_mode_radio,
        )

        self._sync_mode_ui()

    def _is_npm_editor_mode(self):
        return self.npm_editor_mode_radio.isChecked()

    def set_busy(self, busy, text=None):
        super().set_busy(busy, text)
        if not busy:
            self._update_mode_actions()

    def browse_source(self):
        if self._is_multiple_mode():
            self.browse_multiple_sources()
            return
        if self._is_import_mode() or self._is_group_import_mode():
            self.browse_mapping_excel()
            return
        if self._is_cm602_feeder_fix_mode():
            self.browse_mapping_excel()
            return
        if self._is_cm602_program_cm_txt_mode():
            self.browse_cm602_program_source()
            return
        if self._is_cm602_mode():
            self.browse_cm602_source()
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select NPM Machine Export",
            "",
            "NPM Export (*.txt *.TXT *.crb *.CRB);;Text/CRB Files (*.txt *.TXT *.crb *.CRB);;All Files (*)",
        )
        if file_path:
            self.source_picker.set_path(file_path)
            self.source_path = file_path
            self.multi_source_paths = []
            self.mapping_result = None
            self.preview_search_input.clear()
            self.mapping_model.set_records([])
            self.summary_label.setText("0 ROWS")
            self.status_label.setText("Source selected")
            self._update_mode_actions()

    def browse_cm602_source(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select CM602 Feeder/Program File",
            "",
            "CM602/Text Files (*.txt *.TXT);;All Files (*)",
        )
        if file_path:
            self.source_picker.set_path(file_path)
            self.source_path = file_path
            self.multi_source_paths = []
            self.mapping_result = None
            self.preview_search_input.clear()
            self.mapping_model.set_records([])
            self.summary_label.setText("0 ROWS")
            self.status_label.setText("CM602 source selected")
            self._update_mode_actions()

    def browse_cm602_program_source(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select CM602 Program File",
            "",
            "CM602 Program Files (*);;Text Files (*.txt *.TXT)",
        )
        if file_path:
            self.source_picker.set_path(file_path)
            self.source_path = file_path
            self.multi_source_paths = []
            self.mapping_result = None
            self.preview_search_input.clear()
            self.mapping_model.set_records([])
            self.summary_label.setText("Source selected")
            self.status_label.setText("CM602 program selected")
            self._update_mode_actions()

    def browse_mapping_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Feeder Mapping Excel",
            "",
            "Excel Workbook (*.xlsx *.XLSX *.xlsm *.XLSM);;All Files (*)",
        )
        if file_path:
            self.source_picker.set_path(file_path)
            self.source_path = file_path
            self.multi_source_paths = []
            self.mapping_result = None
            self.preview_search_input.clear()
            self.mapping_model.set_records([])
            self.summary_label.setText("Mapping selected")
            self.status_label.setText("Mapping selected")
            self._update_mode_actions()

    def browse_reference_folder(self):
        if self._is_import_mode():
            self.browse_template_npm_program()
            return

        folder_path = QFileDialog.getExistingDirectory(self, "Select Reference Feeder Folder")
        if folder_path:
            self.reference_picker.set_path(folder_path)
            self.status_label.setText("Reference selected")
            self._update_mode_actions()

    def browse_template_npm_program(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select NPM Program Template",
            "",
            "NPM Program (*.crb *.CRB *.txt *.TXT);;Text/CRB Files (*.txt *.TXT *.crb *.CRB);;All Files (*)",
        )
        if file_path:
            self.reference_picker.set_path(file_path)
            self.status_label.setText("Template selected")
            self._update_mode_actions()

    def browse_multiple_sources(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select NPM/CM602 Feeder Files",
            "",
            "NPM/CM602 Files (*.txt *.TXT *.crb *.CRB);;Text/CRB Files (*.txt *.TXT *.crb *.CRB);;All Files (*)",
        )
        if file_paths:
            self.multi_source_paths = file_paths
            self.source_path = ""
            self.mapping_result = None
            self.preview_search_input.clear()
            self.mapping_model.set_records([])
            self.source_picker.set_path("\n".join(file_paths), self._multi_source_display_text(file_paths))
            self.summary_label.setText(f"{len(file_paths)} FILES")
            self.status_label.setText("Sources selected")
            self._update_mode_actions()

    def preview_mapping(self):
        if self._is_multiple_mode() or self._is_import_mode() or self._is_group_import_mode() or self._is_cm602_program_cm_txt_mode() or self._is_cm602_feeder_fix_mode():
            QMessageBox.information(self, "Preview Mapping", "Preview table hanya tersedia untuk mode single feeder file.")
            return

        source_path = self.source_picker.path()
        if not source_path:
            source_label = "File feeder/program CM602" if self._is_cm602_mode() else "File export mesin NPM"
            QMessageBox.warning(self, "Input belum lengkap", f"{source_label} belum dipilih.")
            return
        self.source_path = source_path
        loader = feeder_mapping_service.load_cm602_feeder_mapping if self._is_cm602_mode() else feeder_mapping_service.load_feeder_mapping
        self.run_worker(
            lambda path=source_path, fn=loader: fn(path),
            self._on_mapping_loaded,
            "Loading feeder mapping...",
        )

    def _on_mapping_loaded(self, result):
        self.mapping_result = result
        self.apply_preview_search()
        self.status_label.setText(f"Loaded: {result.source_file}")
        self.generate_btn.setEnabled(True)

    def apply_preview_search(self, *_):
        records = self.mapping_result.records if self.mapping_result is not None else []
        query = self.preview_search_input.text().strip().lower()
        if not query:
            filtered_records = records
        else:
            tokens = [token for token in query.split() if token]
            filtered_records = [
                record
                for record in records
                if all(token in self._preview_search_text(record) for token in tokens)
            ]

        self.mapping_model.set_records(filtered_records)
        self._update_summary(len(filtered_records), len(records))

    def generate_excel(self):
        if self._is_group_import_mode_khusus_sub():
            self.generate_npm_group_import_files_khusus_sub()
            return
        if self._is_import_mode_khusus_sub():
            self.generate_npm_import_file_khusus_sub()
            return
        if self._is_group_import_mode_line8():
            self.generate_npm_group_import_files_line8()
            return
        if self._is_import_mode_line8():
            self.generate_npm_import_file_line8()
            return
        if self._is_group_import_mode():
            self.generate_npm_group_import_files()
            return
        if self._is_import_mode():
            self.generate_npm_import_file()
            return
        if self._is_cm602_program_cm_txt_mode():
            self.generate_cm602_program_cm_txt()
            return
        if self._is_cm602_feeder_fix_mode():
            self.generate_cm602_feeder_fix_file()
            return
        if self._is_multiple_mode():
            self.generate_multiple_excel()
            return

        source_path = self.source_picker.path()
        if not source_path:
            source_label = "File feeder/program CM602" if self._is_cm602_mode() else "File export mesin NPM"
            QMessageBox.warning(self, "Input belum lengkap", f"{source_label} belum dipilih.")
            return

        suggested_name = (
            feeder_mapping_service.suggest_cm602_output_name(source_path)
            if self._is_cm602_mode()
            else feeder_mapping_service.suggest_output_name(source_path)
        )
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Feeder Mapping Excel",
            suggested_name,
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        if self.mapping_result is not None:
            self.run_worker(
                lambda records=self.mapping_result.records, path=output_path: feeder_mapping_service.export_feeder_mapping(records, path),
                self._on_generate_done,
                "Generating feeder mapping Excel...",
            )
            return

        generator = (
            feeder_mapping_service.generate_cm602_feeder_mapping_excel
            if self._is_cm602_mode()
            else feeder_mapping_service.generate_feeder_mapping_excel
        )
        self.run_worker(
            lambda src=source_path, out=output_path, fn=generator: fn(src, out),
            self._on_generate_from_source_done,
            "Generating feeder mapping Excel...",
        )

    def generate_multiple_excel(self):
        source_paths = list(self.multi_source_paths)
        if not source_paths:
            QMessageBox.warning(self, "Input belum lengkap", "File feeder/program NPM atau CM602 belum dipilih.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Multiple Feeder Mapping Excel",
            feeder_mapping_service.suggest_multiple_output_name(source_paths),
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        self.run_worker(
            lambda paths=source_paths, out=output_path: feeder_mapping_service.generate_multiple_feeder_mapping_excel(paths, out),
            self._on_generate_multiple_done,
            "Generating multiple feeder mapping Excel...",
        )

    def generate_cm602_program_cm_txt(self):
        source_path = self.source_picker.path()
        if not source_path:
            QMessageBox.warning(self, "Input belum lengkap", "File program CM602 belum dipilih.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save CM.txt",
            feeder_mapping_service.suggest_cm602_program_cm_txt_output_name(source_path),
            "Text File (*.txt)",
        )
        if not output_path:
            return

        self.run_worker(
            lambda src=source_path, out=output_path: feeder_mapping_service.generate_cm602_program_cm_txt(src, out),
            self._on_generate_cm602_program_cm_txt_done,
            "Converting CM602 program file to CM.txt...",
        )

    def generate_cm602_feeder_fix_file(self):
        mapping_path = self.source_picker.path()
        if not mapping_path:
            QMessageBox.warning(self, "Input belum lengkap", "Excel CM602 feeder mapping belum dipilih.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save CM602 FeederFix TXT",
            feeder_mapping_service.suggest_cm602_feeder_fix_output_name(mapping_path),
            "Text File (*.txt)",
        )
        if not output_path:
            return

        self.run_worker(
            lambda mapping=mapping_path, out=output_path: feeder_mapping_service.generate_cm602_feeder_fix_import_file(mapping, out),
            self._on_generate_cm602_feeder_fix_done,
            "Converting Excel to CM602 FeederFix TXT...",
        )

    def generate_npm_import_file_line8(self):
        mapping_path = self.source_picker.path()
        template_path = self.reference_picker.path()
        if not mapping_path:
            QMessageBox.warning(self, "Input belum lengkap", "Feeder Mapping Excel (Line 8) belum dipilih.")
            return
        if not template_path:
            template_path = ""

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save NPM Feeder Import TXT (Line 8)",
            feeder_mapping_service.suggest_npm_import_output_name(mapping_path),
            "Text File (*.txt)",
        )
        if not output_path:
            return

        self.run_worker(
            lambda mapping=mapping_path, template=template_path, out=output_path: feeder_mapping_service.generate_npm_feeder_import_file_line8(mapping, template, out),
            self._on_generate_npm_import_done,
            "Converting Line 8 feeder mapping to NPM TXT...",
        )

    def generate_npm_group_import_files_line8(self):
        mapping_path = self.source_picker.path()
        template_path = self.reference_picker.path()
        if not mapping_path:
            QMessageBox.warning(self, "Input belum lengkap", "File Fix Feeder Group Excel (Line 8) belum dipilih.")
            return
        if not template_path:
            template_path = ""

        output_dir = QFileDialog.getExistingDirectory(
            self,
            "Pilih Folder Penyimpanan Output NPM TXT (Line 8)",
            str(Path(mapping_path).parent),
        )
        if not output_dir:
            return

        self.run_worker(
            lambda mapping=mapping_path, template=template_path, out=output_dir: feeder_mapping_service.generate_npm_feeder_import_batch_from_groups_line8(mapping, template, out),
            self._on_generate_npm_group_import_done,
            "Converting Line 8 Fix Feeder Groups to NPM TXT...",
        )

    def generate_npm_import_file(self):
        mapping_path = self.source_picker.path()
        template_path = self.reference_picker.path()
        if not mapping_path:
            QMessageBox.warning(self, "Input belum lengkap", "Feeder Mapping Excel belum dipilih.")
            return
        if not template_path:
            template_path = ""

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save NPM Feeder Import TXT",
            feeder_mapping_service.suggest_npm_import_output_name(mapping_path),
            "Text File (*.txt)",
        )
        if not output_path:
            return

        self.run_worker(
            lambda mapping=mapping_path, template=template_path, out=output_path: feeder_mapping_service.generate_npm_feeder_import_file(mapping, template, out),
            self._on_generate_npm_import_done,
            "Converting feeder mapping to NPM TXT...",
        )

    def generate_npm_import_file_khusus_sub(self):
        mapping_path = self.source_picker.path()
        template_path = self.reference_picker.path()
        if not mapping_path:
            QMessageBox.warning(self, "Input belum lengkap", "Feeder Mapping Excel (Khusus Sub) belum dipilih.")
            return
        if not template_path:
            template_path = ""

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save NPM Feeder Import TXT (Khusus Sub)",
            feeder_mapping_service.suggest_npm_import_output_name(mapping_path),
            "Text File (*.txt)",
        )
        if not output_path:
            return

        self.run_worker(
            lambda mapping=mapping_path, template=template_path, out=output_path: feeder_mapping_service.generate_npm_feeder_import_file_khusus_sub(mapping, template, out),
            self._on_generate_npm_import_done,
            "Converting feeder mapping (Khusus Sub) to NPM TXT...",
        )

    def generate_npm_group_import_files_khusus_sub(self):
        mapping_path = self.source_picker.path()
        template_path = self.reference_picker.path()
        if not mapping_path:
            QMessageBox.warning(self, "Input belum lengkap", "File Fix Feeder Group Excel (Khusus Sub) belum dipilih.")
            return
        if not template_path:
            template_path = ""

        output_dir = QFileDialog.getExistingDirectory(
            self,
            "Pilih Folder Penyimpanan Output NPM TXT (Khusus Sub)",
            str(Path(mapping_path).parent),
        )
        if not output_dir:
            return

        self.run_worker(
            lambda mapping=mapping_path, template=template_path, out=output_dir: feeder_mapping_service.generate_npm_feeder_import_batch_from_groups_khusus_sub(mapping, template, out),
            self._on_generate_npm_group_import_done,
            "Converting Fix Feeder Groups (Khusus Sub) to NPM TXT...",
        )

    def generate_npm_group_import_files(self):
        mapping_path = self.source_picker.path()
        template_path = self.reference_picker.path()
        if not mapping_path:
            QMessageBox.warning(self, "Input belum lengkap", "File Fix Feeder Group Excel belum dipilih.")
            return
        if not template_path:
            template_path = ""

        output_dir = QFileDialog.getExistingDirectory(
            self,
            "Pilih Folder Penyimpanan Output NPM TXT",
            str(Path(mapping_path).parent),
        )
        if not output_dir:
            return

        self.run_worker(
            lambda mapping=mapping_path, template=template_path, out=output_dir: feeder_mapping_service.generate_npm_feeder_import_batch_from_groups(mapping, template, out),
            self._on_generate_npm_group_import_done,
            "Converting Fix Feeder Groups to NPM TXT...",
        )

    def _on_generate_npm_group_import_done(self, result):
        output_dir_name = Path(result.output_dir).name
        sub_count = getattr(result, "total_substitute_files", 0)
        total_files = result.total_groups + sub_count
        self.status_label.setText(f"Saved: {total_files} files in {output_dir_name}")
        self.status_label.setToolTip(result.output_dir)
        self.summary_label.setText(f"{result.total_groups} GROUPS | {total_files} FILES | {sub_count} SUB")

        sub_info = f"\nJumlah file _SUB.txt (Substitute): {sub_count}" if sub_count > 0 else ""
        QMessageBox.information(
            self,
            "Fix Feeder Group to NPM TXT",
            (
                f"Jumlah Fix Feeder Group di Excel: {result.total_groups}\n"
                f"Jumlah file TXT utama berhasil dibuat: {result.successful_groups}"
                f"{sub_info}\n\n"
                f"Tersimpan di folder:\n{result.output_dir}"
            ),
        )

    def _on_generate_from_source_done(self, payload):
        result, output_path = payload
        self._on_mapping_loaded(result)
        self._on_generate_done(output_path)

    def _on_generate_multiple_done(self, result):
        output_name = Path(result.output_path).name
        self.status_label.setText(f"Saved: {output_name}")
        self.status_label.setToolTip(result.output_path)
        self.summary_label.setText(f"{result.row_count} ROWS | {result.source_count} FILES | {result.part_count} PARTS")
        QMessageBox.information(
            self,
            "Generate Multiple Feeder Mapping",
            (
                f"Files: {result.source_count}\n"
                f"Rows: {result.row_count}\n"
                f"Unique parts: {result.part_count}\n"
                f"Summary sheet: Summary\n"
                f"File: {result.output_path}"
            ),
        )

    def _on_generate_npm_import_done(self, result):
        output_name = Path(result.output_path).name
        self.status_label.setText(f"Saved: {output_name}")
        self.status_label.setToolTip(result.output_path)
        self.summary_label.setText(f"{result.assignment_count}/{result.mapping_row_count} ROWS | {result.assigned_part_count} PARTS")

        extra = []
        if result.missing_part_rows:
            extra.append("Komponen tidak ada di template")
        if result.missing_location_rows:
            extra.append("Lokasi/Slot tidak ditemukan di template")
        if result.missing_feeder_rows:
            extra.append("Tipe feeder tidak cocok/tidak ditemukan")
        if result.conflict_rows:
            extra.append("Bentrok dengan komponen lain di slot yang sama")
        if result.duplicate_rows:
            extra.append("Baris duplikat diabaikan")
        
        failed_count = result.mapping_row_count - result.assignment_count
        alasan_text = f"\nAlasan gagal: {', '.join(extra)}" if failed_count > 0 and extra else ""

        QMessageBox.information(
            self,
            "Excel to NPM Feeder TXT",
            (
                f"Jumlah komponen di Excel: {result.mapping_row_count}\n"
                f"Yang berhasil masuk: {result.assignment_count}\n"
                f"Gagal masuk: {failed_count}{alasan_text}\n\n"
                f"Tersimpan di:\n{result.output_path}"
            ),
        )

    def _on_generate_cm602_program_cm_txt_done(self, result):
        output_name = Path(result.output_path).name
        self.status_label.setText(f"Saved: {output_name}")
        self.status_label.setToolTip(result.output_path)
        self.summary_label.setText(f"{result.row_count} ROWS | {result.part_count} PARTS")

        QMessageBox.information(
            self,
            "CM602 Program File Converter to CM.txt",
            (
                f"Source: {result.source_file}\n"
                f"Rows: {result.row_count}\n"
                f"Unique parts: {result.part_count}\n"
                f"Board: {result.board_x:.3f} x {result.board_y:.3f}\n"
                f"File: {result.output_path}"
            ),
        )

    def _on_generate_cm602_feeder_fix_done(self, result):
        output_name = Path(result.output_path).name
        self.status_label.setText(f"Saved: {output_name}")
        self.status_label.setToolTip(result.output_path)
        self.summary_label.setText(f"{result.assignment_count}/{result.mapping_row_count} ROWS | {result.slot_count} SLOTS | {result.part_count} PARTS")

        extra = []
        if result.duplicate_rows:
            extra.append(f"Duplicate rows skipped: {len(result.duplicate_rows)}")
        if result.default_feeder_rows:
            extra.append(f"Default feeder used: {len(result.default_feeder_rows)}")
        extra_text = "\n" + "\n".join(extra) if extra else ""

        QMessageBox.information(
            self,
            "Excel to CM602 FeederFix TXT",
            (
                f"Mapping: {result.mapping_file}\n"
                f"Mapping rows: {result.mapping_row_count}\n"
                f"Assigned rows: {result.assignment_count}\n"
                f"Slots: {result.slot_count}\n"
                f"Unique parts: {result.part_count}{extra_text}\n"
                f"File: {result.output_path}"
            ),
        )

    def _on_generate_done(self, output_path):
        output_name = Path(output_path).name
        self.status_label.setText(f"Saved: {output_name}")
        self.status_label.setToolTip(output_path)
        QMessageBox.information(self, "Generate Feeder Mapping", f"Exported to:\n{output_path}")

    def browse_source(self):
        if self._is_npm_editor_mode():
            self.browse_npm_editor_source()
            return
        if self._is_multiple_mode():
            self.browse_multiple_sources()
            return
        if self._is_import_mode() or self._is_group_import_mode():
            self.browse_mapping_excel()
            return
        if self._is_cm602_feeder_fix_mode():
            self.browse_mapping_excel()
            return
        if self._is_cm602_program_cm_txt_mode():
            self.browse_cm602_program_source()
            return
        if self._is_cm602_mode():
            self.browse_cm602_source()
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select NPM Machine Export",
            "",
            "NPM Export (*.txt *.TXT *.crb *.CRB);;Text/CRB Files (*.txt *.TXT *.crb *.CRB);;All Files (*)",
        )
        if file_path:
            self.source_picker.set_path(file_path)
            self.source_path = file_path
            self.multi_source_paths = []
            self.mapping_result = None
            self.preview_search_input.clear()
            self.mapping_model.set_records([])
            self.summary_label.setText("0 ROWS")
            self.status_label.setText("Source selected")
            self._update_mode_actions()

    def browse_npm_editor_source(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select NPM Feeder TXT File",
            "",
            "NPM TXT Files (*.txt *.TXT);;All Files (*)",
        )
        if file_path:
            self.source_picker.set_path(file_path)
            self.source_path = file_path
            self.multi_source_paths = []
            try:
                self.npm_doc = npm_txt_editor_service.load_npm_txt(file_path)
                self._reload_npm_editor_table()
                self.status_label.setText("File TXT berhasil dimuat ke Editor")
            except Exception as e:
                QMessageBox.critical(self, "Error Load TXT", f"Gagal membaca file NPM TXT:\n{e}")
            self._update_mode_actions()

    def _reload_npm_editor_table(self):
        if not self.npm_doc:
            self.mapping_model.set_records([])
            self.summary_label.setText("0 SLOTS")
            return
        slots = self.npm_doc.get_slots()
        self.npm_doc_slots = slots
        records = []
        for s in slots:
            loc = s["location_code"]
            table_val = ""
            slot_val = ""
            pos_val = ""
            m = re.match(r'^\[(\d+)\](\d+)(?:-(\d+))?([LRlr])?$', loc)
            if m:
                table_val = int(m.group(1))
                slot_val = int(m.group(2))
                pos_val = m.group(4).upper() if m.group(4) else "Single"
            records.append({
                "table": table_val,
                "slot": slot_val,
                "position": pos_val,
                "location_code": loc,
                "part_number": s["part_number"],
                "feeder_id": s["feeder_id"],
                "pu_code": s["pu_code"],
            })
        self.mapping_model.set_records(records)
        self.summary_label.setText(f"{len(records)} SLOTS")

    def editor_add_slot(self):
        if not self.npm_doc:
            QMessageBox.warning(self, "Peringatan", "Silakan pilih file NPM Feeder TXT terlebih dahulu.")
            return
        dlg = NpmTxtSlotEditDialog(self)
        if dlg.exec() == QDialog.Accepted:
            data = dlg.get_data()
            if not data["part_number"]:
                QMessageBox.warning(self, "Peringatan", "Part Number tidak boleh kosong.")
                return
            self.npm_doc.add_slot(
                table=data["table"],
                slot=data["slot"],
                pos=data["pos"],
                part_number=data["part_number"],
                feeder_id=data["feeder_id"],
            )
            self._reload_npm_editor_table()
            self.status_label.setText(f"Slot [{data['table']}]{data['slot']:02d}{data['pos']} berhasil ditambahkan/diupdate.")
            self._update_mode_actions()

    def editor_edit_slot(self):
        if not self.npm_doc:
            QMessageBox.warning(self, "Peringatan", "Silakan pilih file NPM Feeder TXT terlebih dahulu.")
            return
        selected_indexes = self.mapping_table.selectionModel().selectedRows()
        if not selected_indexes:
            QMessageBox.warning(self, "Peringatan", "Pilih slot feeder yang ingin diedit dari tabel terlebih dahulu.")
            return
        row_idx = selected_indexes[0].row()
        records = self.mapping_model.records
        if row_idx < 0 or row_idx >= len(records):
            return
        rec = records[row_idx]
        table_val = rec.get("table", 1) or 1
        slot_val = rec.get("slot", 1) or 1
        pos_val = rec.get("position", "")
        if pos_val == "Single":
            pos_val = ""
        part_number = rec.get("part_number", "")
        feeder_id = rec.get("feeder_id", "")

        dlg = NpmTxtSlotEditDialog(
            self,
            table=table_val,
            slot=slot_val,
            pos=pos_val,
            part_number=part_number,
            feeder_id=feeder_id,
        )
        if dlg.exec() == QDialog.Accepted:
            data = dlg.get_data()
            if not data["part_number"]:
                QMessageBox.warning(self, "Peringatan", "Part Number tidak boleh kosong.")
                return
            self.npm_doc.add_slot(
                table=data["table"],
                slot=data["slot"],
                pos=data["pos"],
                part_number=data["part_number"],
                feeder_id=data["feeder_id"],
            )
            self._reload_npm_editor_table()
            self.status_label.setText(f"Slot [{data['table']}]{data['slot']:02d}{data['pos']} berhasil diperbarui.")
            self._update_mode_actions()

    def editor_delete_slot(self):
        if not self.npm_doc:
            QMessageBox.warning(self, "Peringatan", "Silakan pilih file NPM Feeder TXT terlebih dahulu.")
            return
        selected_indexes = self.mapping_table.selectionModel().selectedRows()
        if not selected_indexes:
            QMessageBox.warning(self, "Peringatan", "Pilih slot feeder yang ingin dihapus dari tabel terlebih dahulu.")
            return
        row_idx = selected_indexes[0].row()
        records = self.mapping_model.records
        if row_idx < 0 or row_idx >= len(records):
            return
        rec = records[row_idx]
        pu_code = rec.get("pu_code")
        loc_code = rec.get("location_code")
        reply = QMessageBox.question(
            self,
            "Konfirmasi Hapus Slot",
            f"Apakah Anda yakin ingin menghapus slot {loc_code} ({rec.get('part_number')})?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self.npm_doc.delete_slot(pu_code)
            self._reload_npm_editor_table()
            self.status_label.setText(f"Slot {loc_code} berhasil dihapus.")
            self._update_mode_actions()

    def editor_save_file(self):
        if not self.npm_doc:
            QMessageBox.warning(self, "Peringatan", "Silakan pilih dan edit file NPM Feeder TXT terlebih dahulu.")
            return
        default_name = "EDITED_" + Path(self.source_path).name if self.source_path else "NPM_FEEDER_EDITED.txt"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Simpan / Export NPM Feeder TXT",
            default_name,
            "NPM TXT Files (*.txt *.TXT);;All Files (*)",
        )
        if file_path:
            try:
                npm_txt_editor_service.save_npm_txt(self.npm_doc, file_path)
                QMessageBox.information(
                    self,
                    "Berhasil",
                    f"File NPM Feeder TXT berhasil disimpan di:\n{file_path}",
                )
                self.status_label.setText("File berhasil disimpan")
            except Exception as e:
                QMessageBox.critical(self, "Error Simpan File", f"Gagal menyimpan file NPM TXT:\n{e}")

    def browse_cm602_source(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select CM602 Feeder/Program File",
            "",
            "CM602/Text Files (*.txt *.TXT);;All Files (*)",
        )
        if file_path:
            self.source_picker.set_path(file_path)
            self.source_path = file_path
            self.multi_source_paths = []
            self.mapping_result = None
            self.preview_search_input.clear()
            self.mapping_model.set_records([])
            self.summary_label.setText("0 ROWS")
            self.status_label.setText("CM602 source selected")
            self._update_mode_actions()

    def browse_cm602_program_source(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select CM602 Program File",
            "",
            "CM602 Program Files (*);;Text Files (*.txt *.TXT)",
        )
        if file_path:
            self.source_picker.set_path(file_path)
            self.source_path = file_path
            self.multi_source_paths = []
            self.mapping_result = None
            self.preview_search_input.clear()
            self.mapping_model.set_records([])
            self.summary_label.setText("Source selected")
            self.status_label.setText("CM602 program selected")
            self._update_mode_actions()

    def clear_data(self):
        self.mapping_result = None
        self.source_path = ""
        self.multi_source_paths = []
        self.npm_doc = None
        self.npm_doc_slots = []
        self.source_picker.clear()
        self.reference_picker.clear()
        self.balancing_input.setPlainText(feeder_mapping_service.default_balancing_part_numbers_text())
        self.preview_search_input.clear()
        self.mapping_model.set_records([])
        if self._is_npm_editor_mode():
            self.summary_label.setText("0 SLOTS")
        elif self._is_import_mode() or self._is_cm602_feeder_fix_mode():
            self.summary_label.setText("0 PARTS")
        else:
            self.summary_label.setText("0 FILES" if self._is_multiple_mode() else "0 ROWS")
        self.status_label.setText("")
        self.status_label.setToolTip("")
        self._update_mode_actions()

    def _preview_search_text(self, record):
        return " ".join(
            str(record.get(key, ""))
            for key in ("table", "slot", "position", "location_code", "part_number")
        ).lower()

    def _update_summary(self, visible_count, total_count):
        if self._is_npm_editor_mode():
            if self.npm_doc is None:
                self.summary_label.setText("0 SLOTS")
                return
            summary = f"{total_count} SLOTS"
            if visible_count != total_count:
                summary = f"{visible_count}/{summary}"
            self.summary_label.setText(summary)
            return

        if self.mapping_result is None:
            self.summary_label.setText("0 ROWS")
            return

        summary = f"{total_count} ROWS | {self.mapping_result.table_count} TABLES | {self.mapping_result.part_count} PARTS"
        if visible_count != total_count:
            summary = f"{visible_count}/{summary}"
        self.summary_label.setText(summary)

    def _is_multiple_mode(self):
        return self.multiple_mode_radio.isChecked()

    def _is_cm602_mode(self):
        return self.cm602_mode_radio.isChecked()

    def _is_cm602_program_cm_txt_mode(self):
        return self.cm602_program_cm_txt_mode_radio.isChecked()

    def _is_cm602_feeder_fix_mode(self):
        return self.cm602_feeder_fix_mode_radio.isChecked()

    def _is_import_mode(self):
        return self.import_mode_radio.isChecked()

    def _is_import_mode_line8(self):
        return self.import_mode_line8_radio.isChecked()

    def _is_import_mode_khusus_sub(self):
        return self.import_mode_khusus_sub_radio.isChecked()

    def _is_group_import_mode(self):
        return self.group_import_mode_radio.isChecked()

    def _is_group_import_mode_line8(self):
        return self.group_import_mode_line8_radio.isChecked()

    def _is_group_import_mode_khusus_sub(self):
        return self.group_import_mode_khusus_sub_radio.isChecked()

    def _is_npm_editor_mode(self):
        return self.npm_editor_mode_radio.isChecked()

    def _sync_mode_ui(self):
        is_multiple = self._is_multiple_mode()
        is_cm602 = self._is_cm602_mode()
        is_cm602_program_cm_txt = self._is_cm602_program_cm_txt_mode()
        is_cm602_feeder_fix = self._is_cm602_feeder_fix_mode()
        is_import = self._is_import_mode() or self._is_import_mode_line8() or self._is_import_mode_khusus_sub()
        is_group_import = self._is_group_import_mode() or self._is_group_import_mode_line8() or self._is_group_import_mode_khusus_sub()
        is_npm_editor = self._is_npm_editor_mode()

        if is_npm_editor:
            self.source_picker.label.setText("NPM Feeder TXT File (.txt):")
            self.source_picker.button.setText("Browse")
        elif is_group_import:
            self.source_picker.label.setText("Fix Feeder Group Excel (.xlsx):")
            self.reference_picker.label.setText("NPM Program Template (Optional, kosongkan untuk auto Line 8):" if self._is_group_import_mode_line8() else "NPM Program Template (Optional, kosongkan untuk auto):")
            self.source_picker.button.setText("Browse")
            self.reference_picker.button.setText("Browse")
        elif is_import:
            self.source_picker.label.setText("Feeder Mapping Excel (.xlsx):")
            self.reference_picker.label.setText("NPM Program Template (Optional, kosongkan untuk auto Line 8):" if self._is_import_mode_line8() else "NPM Program Template (Optional, kosongkan untuk auto):")
            self.source_picker.button.setText("Browse")
            self.reference_picker.button.setText("Browse")
        elif is_cm602_feeder_fix:
            self.source_picker.label.setText("CM602 Feeder Mapping Excel (.xlsx):")
            self.source_picker.button.setText("Browse")
        elif is_cm602_program_cm_txt:
            self.source_picker.label.setText("CM602 Program File:")
            self.source_picker.button.setText("Browse")
        elif is_cm602:
            self.source_picker.label.setText("CM602 Feeder/Program File:")
            self.source_picker.button.setText("Browse")
        else:
            self.source_picker.label.setText("NPM/CM602 Feeder Files:" if is_multiple else "NPM Machine Export (.txt/.crb):")
            self.source_picker.button.setText("Browse Files" if is_multiple else "Browse")
        self.reference_picker.setVisible(is_import or is_group_import)
        self.balancing_label.setVisible(False)
        self.balancing_input.setVisible(False)
        self.mapping_result = None
        self.npm_doc = None
        self.npm_doc_slots = []
        self.source_path = ""
        self.multi_source_paths = []
        self.source_picker.clear()
        self.reference_picker.clear()
        self.preview_search_input.clear()
        self.mapping_model.set_records([])
        if is_npm_editor:
            self.summary_label.setText("0 SLOTS")
        elif is_import or is_group_import or is_cm602_feeder_fix:
            self.summary_label.setText("0 GROUPS" if is_group_import else "0 PARTS")
        else:
            self.summary_label.setText("0 FILES" if is_multiple else "0 ROWS")
        self.status_label.setText("")
        self.status_label.setToolTip("")
        self._update_mode_actions()

    def _update_mode_actions(self):
        is_multiple = self._is_multiple_mode()
        is_import = self._is_import_mode()
        is_group_import = self._is_group_import_mode()
        is_cm602_program_cm_txt = self._is_cm602_program_cm_txt_mode()
        is_cm602_feeder_fix = self._is_cm602_feeder_fix_mode()
        is_npm_editor = self._is_npm_editor_mode()

        self.preview_btn.setVisible(not is_npm_editor)
        self.generate_btn.setVisible(not is_npm_editor)
        self.editor_add_btn.setVisible(is_npm_editor)
        self.editor_edit_btn.setVisible(is_npm_editor)
        self.editor_delete_btn.setVisible(is_npm_editor)
        self.editor_save_btn.setVisible(is_npm_editor)

        if is_npm_editor:
            has_doc = bool(self.npm_doc)
            self.editor_add_btn.setEnabled(has_doc)
            self.editor_edit_btn.setEnabled(has_doc)
            self.editor_delete_btn.setEnabled(has_doc)
            self.editor_save_btn.setEnabled(has_doc)
            self.preview_search_input.setEnabled(has_doc)
        else:
            self.preview_btn.setEnabled(not is_multiple and not is_import and not is_group_import and not is_cm602_program_cm_txt and not is_cm602_feeder_fix)
            self.preview_search_input.setEnabled(not is_multiple and not is_import and not is_group_import and not is_cm602_program_cm_txt and not is_cm602_feeder_fix)
            if is_group_import:
                self.generate_btn.setText("Generate Group NPM TXTs")
                self.generate_btn.setEnabled(bool(self.source_picker.path()))
            elif is_import:
                self.generate_btn.setText("Generate Feeder TXT")
                self.generate_btn.setEnabled(bool(self.source_picker.path()))
            elif is_cm602_program_cm_txt:
                self.generate_btn.setText("Generate CM.txt")
                self.generate_btn.setEnabled(bool(self.source_picker.path()))
            elif is_cm602_feeder_fix:
                self.generate_btn.setText("Generate FeederFix TXT")
                self.generate_btn.setEnabled(bool(self.source_picker.path()))
            else:
                self.generate_btn.setText("Generate Workbook" if is_multiple else "Generate Excel")
                self.generate_btn.setEnabled(bool(self.multi_source_paths) if is_multiple else bool(self.source_picker.path()))

    def _multi_source_display_text(self, file_paths):
        file_names = [Path(path).name for path in file_paths]
        preview = ", ".join(file_names[:3])
        if len(file_names) > 3:
            preview = f"{preview}, +{len(file_names) - 3} more"
        return f"{len(file_names)} files selected: {preview}"
