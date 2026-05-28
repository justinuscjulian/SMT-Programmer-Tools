from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
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

        source_card = Card()
        source_title = QLabel("Source File")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)
        self.source_picker = FilePicker("NPM Machine Export (.txt):")
        self.source_picker.browse_requested.connect(self.browse_source)
        source_card.layout.addWidget(self.source_picker)
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
        table_title = QLabel("Detailed Feeder Setup")
        table_title.setObjectName("SectionTitle")
        table_card.layout.addWidget(table_title)
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
            self.clear_btn,
            self.source_picker.button,
        )

    def browse_source(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select NPM Machine Export",
            "",
            "NPM Export (*.txt);;Text Files (*.txt);;All Files (*)",
        )
        if file_path:
            self.source_picker.set_path(file_path)
            self.mapping_result = None
            self.mapping_model.set_records([])
            self.generate_btn.setEnabled(True)
            self.summary_label.setText("0 ROWS")
            self.status_label.setText("Source selected")

    def preview_mapping(self):
        source_path = self.source_picker.path()
        if not source_path:
            QMessageBox.warning(self, "Input belum lengkap", "File export mesin NPM belum dipilih.")
            return
        self.source_path = source_path
        self.run_worker(
            lambda path=source_path: feeder_mapping_service.load_feeder_mapping(path),
            self._on_mapping_loaded,
            "Loading feeder mapping...",
        )

    def _on_mapping_loaded(self, result):
        self.mapping_result = result
        self.mapping_model.set_records(result.records)
        self.summary_label.setText(f"{result.row_count} ROWS | {result.table_count} TABLES | {result.part_count} PARTS")
        self.status_label.setText(f"Loaded: {result.source_file}")
        self.generate_btn.setEnabled(True)

    def generate_excel(self):
        source_path = self.source_picker.path()
        if not source_path:
            QMessageBox.warning(self, "Input belum lengkap", "File export mesin NPM belum dipilih.")
            return

        suggested_name = feeder_mapping_service.suggest_output_name(source_path)
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

        self.run_worker(
            lambda src=source_path, out=output_path: feeder_mapping_service.generate_feeder_mapping_excel(src, out),
            self._on_generate_from_source_done,
            "Generating feeder mapping Excel...",
        )

    def _on_generate_from_source_done(self, payload):
        result, output_path = payload
        self._on_mapping_loaded(result)
        self._on_generate_done(output_path)

    def _on_generate_done(self, output_path):
        output_name = Path(output_path).name
        self.status_label.setText(f"Saved: {output_name}")
        self.status_label.setToolTip(output_path)
        QMessageBox.information(self, "Generate Feeder Mapping", f"Exported to:\n{output_path}")

    def clear_data(self):
        self.mapping_result = None
        self.source_path = ""
        self.source_picker.clear()
        self.mapping_model.set_records([])
        self.generate_btn.setEnabled(False)
        self.summary_label.setText("0 ROWS")
        self.status_label.setText("")
        self.status_label.setToolTip("")
