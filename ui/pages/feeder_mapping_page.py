from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QTableView,
    QVBoxLayout,
)

from models.table_model import ColumnSpec, RecordTableModel
from services import feeder_mapping_service
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu


class FeederMappingPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.mapping_result = None
        self.source_path = ""
        self.multi_source_paths = []
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
        mode_row = QHBoxLayout()
        self.single_mode_radio = QRadioButton("Single Feeder File")
        self.multiple_mode_radio = QRadioButton("Multiple Feeder Files")
        self.cm602_mode_radio = QRadioButton("CM602")
        self.cm602_program_cm_txt_mode_radio = QRadioButton("CM602 Program File Converter to CM.txt")
        self.cm602_feeder_fix_mode_radio = QRadioButton("Excel to CM602 FeederFix TXT")
        self.import_mode_radio = QRadioButton("Excel to NPM Feeder TXT")
        self.single_mode_radio.setChecked(True)
        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.single_mode_radio)
        self.mode_group.addButton(self.multiple_mode_radio)
        self.mode_group.addButton(self.cm602_mode_radio)
        self.mode_group.addButton(self.cm602_program_cm_txt_mode_radio)
        self.mode_group.addButton(self.cm602_feeder_fix_mode_radio)
        self.mode_group.addButton(self.import_mode_radio)
        self.single_mode_radio.toggled.connect(self._sync_mode_ui)
        self.multiple_mode_radio.toggled.connect(self._sync_mode_ui)
        self.cm602_mode_radio.toggled.connect(self._sync_mode_ui)
        self.cm602_program_cm_txt_mode_radio.toggled.connect(self._sync_mode_ui)
        self.cm602_feeder_fix_mode_radio.toggled.connect(self._sync_mode_ui)
        self.import_mode_radio.toggled.connect(self._sync_mode_ui)
        mode_row.addWidget(self.single_mode_radio)
        mode_row.addWidget(self.multiple_mode_radio)
        mode_row.addWidget(self.cm602_mode_radio)
        mode_row.addWidget(self.cm602_program_cm_txt_mode_radio)
        mode_row.addWidget(self.cm602_feeder_fix_mode_radio)
        mode_row.addWidget(self.import_mode_radio)
        mode_row.addStretch(1)
        mode_card.layout.addLayout(mode_row)
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
        action_bar.addWidget(self.preview_btn)
        action_bar.addWidget(self.generate_btn)
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
        )

        self._sync_mode_ui()

    def set_busy(self, busy, text=None):
        super().set_busy(busy, text)
        if not busy:
            self._update_mode_actions()

    def browse_source(self):
        if self._is_multiple_mode():
            self.browse_multiple_sources()
            return
        if self._is_import_mode():
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
        if self._is_multiple_mode() or self._is_import_mode() or self._is_cm602_program_cm_txt_mode() or self._is_cm602_feeder_fix_mode():
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
            extra.append(f"Missing parts: {len(result.missing_part_rows)}")
        if result.missing_location_rows:
            extra.append(f"Missing location: {len(result.missing_location_rows)}")
        if result.missing_feeder_rows:
            extra.append(f"Missing feeder type: {len(result.missing_feeder_rows)}")
        if result.conflict_rows:
            extra.append(f"Conflicts: {len(result.conflict_rows)}")
        if result.duplicate_rows:
            extra.append(f"Duplicate rows skipped: {len(result.duplicate_rows)}")
        extra_text = "\n" + "\n".join(extra) if extra else ""

        QMessageBox.information(
            self,
            "Excel to NPM Feeder TXT",
            (
                f"Mapping: {result.mapping_file}\n"
                f"Template: {result.template_file}\n"
                f"Mapping rows: {result.mapping_row_count}\n"
                f"Assigned rows: {result.assignment_count}\n"
                f"Assigned parts: {result.assigned_part_count}{extra_text}\n"
                f"File: {result.output_path}"
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

    def clear_data(self):
        self.mapping_result = None
        self.source_path = ""
        self.multi_source_paths = []
        self.source_picker.clear()
        self.reference_picker.clear()
        self.balancing_input.setPlainText(feeder_mapping_service.default_balancing_part_numbers_text())
        self.preview_search_input.clear()
        self.mapping_model.set_records([])
        if self._is_import_mode() or self._is_cm602_feeder_fix_mode():
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

    def _sync_mode_ui(self):
        is_multiple = self._is_multiple_mode()
        is_cm602 = self._is_cm602_mode()
        is_cm602_program_cm_txt = self._is_cm602_program_cm_txt_mode()
        is_cm602_feeder_fix = self._is_cm602_feeder_fix_mode()
        is_import = self._is_import_mode()
        if is_import:
            self.source_picker.label.setText("Feeder Mapping Excel (.xlsx):")
            self.reference_picker.label.setText("NPM Program Template (Optional, kosongkan untuk auto):")
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
        self.reference_picker.setVisible(is_import)
        self.balancing_label.setVisible(False)
        self.balancing_input.setVisible(False)
        self.mapping_result = None
        self.source_path = ""
        self.multi_source_paths = []
        self.source_picker.clear()
        self.reference_picker.clear()
        self.preview_search_input.clear()
        self.mapping_model.set_records([])
        if is_import or is_cm602_feeder_fix:
            self.summary_label.setText("0 PARTS")
        else:
            self.summary_label.setText("0 FILES" if is_multiple else "0 ROWS")
        self.status_label.setText("")
        self.status_label.setToolTip("")
        self._update_mode_actions()

    def _update_mode_actions(self):
        is_multiple = self._is_multiple_mode()
        is_import = self._is_import_mode()
        is_cm602_program_cm_txt = self._is_cm602_program_cm_txt_mode()
        is_cm602_feeder_fix = self._is_cm602_feeder_fix_mode()
        self.preview_btn.setEnabled(not is_multiple and not is_import and not is_cm602_program_cm_txt and not is_cm602_feeder_fix)
        self.preview_search_input.setEnabled(not is_multiple and not is_import and not is_cm602_program_cm_txt and not is_cm602_feeder_fix)
        if is_import:
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
