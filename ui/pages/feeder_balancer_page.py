from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTableView,
    QTabWidget,
    QVBoxLayout,
)

from models.table_model import ColumnSpec, RecordTableModel
from services.feeder_balancer_service import (
    BALANCED_COLUMNS,
    DEFAULT_DUPLICATE_MIN_INSERT,
    DEFAULT_MULTI_COPY_MIN_INSERT,
    DUPLICATE_PLAN_COLUMNS,
    FIELD_LABELS,
    MACHINE_MODES,
    MACHINE_NPM_CUSTOM,
    WARNING_COLUMNS,
    ZONE_SUMMARY_COLUMNS,
    FeederBalancerConfig,
    analyze_feeder_balance,
    detect_feeder_balancer_zones,
    export_feeder_balance_result,
    load_feeder_balancer_preview,
    suggest_export_name,
)
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu
from workers.task_runner import TaskWorker


class FeederBalancerPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.preview_result = None
        self.zone_result = None
        self.balance_result = None
        self.mapping_combos = {}
        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Feeder Balancer")
        title.setObjectName("SectionTitle")
        subtitle = QLabel("Auto-balance fixed feeder arrangement from imported SLOT, PART NUMBER, and COMPONENT INSERT data.")
        subtitle.setObjectName("MutedLabel")
        self.summary_label = QLabel("0 ZONES | 0 PARTS")
        self.summary_label.setObjectName("MutedLabel")
        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addWidget(subtitle, 1)
        header.addWidget(self.summary_label)
        header.addWidget(self.status_label)
        root.addLayout(header)

        source_card = Card()
        source_title = QLabel("Source Data")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)

        self.source_picker = FilePicker("Excel/CSV Mapping:")
        self.source_picker.browse_requested.connect(self.browse_source)
        source_card.layout.addWidget(self.source_picker)

        mode_row = QHBoxLayout()
        mode_label = QLabel("Machine Mode:")
        mode_label.setMinimumWidth(150)
        self.machine_combo = QComboBox()
        self.machine_combo.addItems(MACHINE_MODES)
        self.machine_combo.currentTextChanged.connect(self._sync_profile_ui)
        mode_row.addWidget(mode_label)
        mode_row.addWidget(self.machine_combo)
        mode_row.addStretch(1)
        source_card.layout.addLayout(mode_row)

        duplicate_rule_row = QHBoxLayout()
        duplicate_min_label = QLabel("Min Duplicate Insert:")
        duplicate_min_label.setMinimumWidth(150)
        self.duplicate_min_spin = QSpinBox()
        self.duplicate_min_spin.setRange(0, 999999)
        self.duplicate_min_spin.setValue(int(DEFAULT_DUPLICATE_MIN_INSERT))
        self.duplicate_min_spin.valueChanged.connect(self._on_mapping_changed)
        multi_copy_label = QLabel("Multi-copy Insert:")
        multi_copy_label.setMinimumWidth(130)
        self.multi_copy_spin = QSpinBox()
        self.multi_copy_spin.setRange(0, 999999)
        self.multi_copy_spin.setValue(int(DEFAULT_MULTI_COPY_MIN_INSERT))
        self.multi_copy_spin.valueChanged.connect(self._on_mapping_changed)
        duplicate_rule_row.addWidget(duplicate_min_label)
        duplicate_rule_row.addWidget(self.duplicate_min_spin)
        duplicate_rule_row.addSpacing(16)
        duplicate_rule_row.addWidget(multi_copy_label)
        duplicate_rule_row.addWidget(self.multi_copy_spin)
        duplicate_rule_row.addStretch(1)
        source_card.layout.addLayout(duplicate_rule_row)

        self.profile_label = QLabel("NPM Custom Profile (optional constraints):")
        self.profile_label.setObjectName("MutedLabel")
        self.profile_input = QPlainTextEdit()
        self.profile_input.setPlaceholderText(
            "Optional key=value, one per line.\n"
            "Example:\n"
            "max_copy_per_part=3\n"
            "duplicate_min_insert=10\n"
            "multi_copy_min_insert=25\n"
            "reserved_slots=[1]1L, [1]2L"
        )
        self.profile_input.setMaximumHeight(92)
        source_card.layout.addWidget(self.profile_label)
        source_card.layout.addWidget(self.profile_input)
        root.addWidget(source_card)

        mapping_card = Card()
        mapping_header = QHBoxLayout()
        mapping_title = QLabel("Column Mapping")
        mapping_title.setObjectName("SectionTitle")
        self.detected_columns_label = QLabel("No file loaded")
        self.detected_columns_label.setObjectName("MutedLabel")
        mapping_header.addWidget(mapping_title)
        mapping_header.addStretch(1)
        mapping_header.addWidget(self.detected_columns_label)
        mapping_card.layout.addLayout(mapping_header)

        mapping_grid = QGridLayout()
        mapping_grid.setHorizontalSpacing(12)
        mapping_grid.setVerticalSpacing(8)
        for index, (field_name, label_text) in enumerate(FIELD_LABELS.items()):
            label = QLabel(f"{label_text}:")
            label.setMinimumWidth(135)
            combo = QComboBox()
            combo.addItem("Not used")
            combo.currentIndexChanged.connect(self._on_mapping_changed)
            self.mapping_combos[field_name] = combo
            row, column_group = divmod(index, 2)
            mapping_grid.addWidget(label, row, column_group * 2)
            mapping_grid.addWidget(combo, row, column_group * 2 + 1)
        mapping_card.layout.addLayout(mapping_grid)
        root.addWidget(mapping_card)

        action_bar = QHBoxLayout()
        self.preview_btn = QPushButton("Preview Detection")
        self.preview_btn.setObjectName("PrimaryButton")
        self.preview_btn.clicked.connect(self.preview_detection)
        self.generate_btn = QPushButton("Generate Balance")
        self.generate_btn.setObjectName("SuccessButton")
        self.generate_btn.clicked.connect(self.generate_balance)
        self.export_btn = QPushButton("Export Excel")
        self.export_btn.setObjectName("SuccessButton")
        self.export_btn.clicked.connect(self.export_excel)
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_data)
        self.preview_btn.setEnabled(False)
        self.generate_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        action_bar.addWidget(self.preview_btn)
        action_bar.addWidget(self.generate_btn)
        action_bar.addWidget(self.export_btn)
        action_bar.addStretch(1)
        action_bar.addWidget(self.clear_btn)
        root.addLayout(action_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        tabs = QTabWidget()
        tabs.addTab(self._build_detected_columns_tab(), "Detected Columns")
        tabs.addTab(self._build_zone_tab(), "Detected Zones")
        tabs.addTab(self._build_balanced_tab(), "Balanced Feeder")
        tabs.addTab(self._build_summary_tab(), "Summary")
        tabs.addTab(self._build_duplicate_tab(), "Duplicate Plan")
        tabs.addTab(self._build_warning_tab(), "Warnings")
        root.addWidget(tabs, 1)

        self.register_busy_widgets(
            self.source_picker.button,
            self.machine_combo,
            self.profile_input,
            self.preview_btn,
            self.generate_btn,
            self.export_btn,
            self.clear_btn,
            self.duplicate_min_spin,
            self.multi_copy_spin,
            *self.mapping_combos.values(),
        )
        self._sync_profile_ui()

    def _build_detected_columns_tab(self):
        page = Card()
        header = QLabel("Input Preview")
        header.setObjectName("SectionTitle")
        page.layout.addWidget(header)
        self.preview_model = RecordTableModel(theme=self.theme_manager.theme)
        self.register_model(self.preview_model)
        self.preview_table = QTableView()
        configure_table(self.preview_table, self.preview_model, wrap_headers=True)
        install_copy_menu(self.preview_table, self.preview_model)
        page.layout.addWidget(self.preview_table, 1)
        return page

    def _build_zone_tab(self):
        page = Card()
        header = QLabel("Detected Zone / Table / Module")
        header.setObjectName("SectionTitle")
        page.layout.addWidget(header)
        self.zone_model = RecordTableModel(
            [
                ColumnSpec("zone", "Zone", Qt.AlignCenter, 100),
                ColumnSpec("slot_count", "Slot Count", Qt.AlignCenter, 95),
                ColumnSpec("usable_slot_count", "Usable Slots", Qt.AlignCenter, 95),
                ColumnSpec("parsed_zone_count", "Parsed Zone", Qt.AlignCenter, 95),
                ColumnSpec("parsed_slot_number_count", "Parsed Slot No", Qt.AlignCenter, 110),
                ColumnSpec("parsed_side_count", "Parsed Side", Qt.AlignCenter, 95),
                ColumnSpec("parse_status", "Parse Status", Qt.AlignCenter, 110),
                ColumnSpec("sample_slots", "Sample Slots", Qt.AlignLeft, 420),
            ],
            status_key="parse_status",
            theme=self.theme_manager.theme,
        )
        self.register_model(self.zone_model)
        self.zone_table = QTableView()
        configure_table(self.zone_table, self.zone_model, wrap_headers=True)
        install_copy_menu(self.zone_table, self.zone_model)
        page.layout.addWidget(self.zone_table, 1)
        return page

    def _build_balanced_tab(self):
        page = Card()
        header = QLabel("Balanced Feeder Table")
        header.setObjectName("SectionTitle")
        page.layout.addWidget(header)
        self.balanced_model = RecordTableModel(
            [ColumnSpec(key, header, Qt.AlignLeft if key in {"part_number", "assigned_zones", "source_slot", "note"} else Qt.AlignCenter, 150) for key, header in BALANCED_COLUMNS],
            status_key="note",
            theme=self.theme_manager.theme,
        )
        self.register_model(self.balanced_model)
        self.balanced_table = QTableView()
        configure_table(self.balanced_table, self.balanced_model, wrap_headers=True)
        install_copy_menu(self.balanced_table, self.balanced_model)
        page.layout.addWidget(self.balanced_table, 1)
        return page

    def _build_summary_tab(self):
        page = Card()
        header = QLabel("Summary Per Zone")
        header.setObjectName("SectionTitle")
        page.layout.addWidget(header)
        self.summary_model = RecordTableModel(
            [ColumnSpec(key, header, Qt.AlignCenter, 145) for key, header in ZONE_SUMMARY_COLUMNS],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.summary_model)
        self.summary_table = QTableView()
        configure_table(self.summary_table, self.summary_model, wrap_headers=True)
        install_copy_menu(self.summary_table, self.summary_model)
        page.layout.addWidget(self.summary_table, 1)
        return page

    def _build_duplicate_tab(self):
        page = Card()
        header = QLabel("Duplicate Feeder Plan")
        header.setObjectName("SectionTitle")
        page.layout.addWidget(header)
        self.duplicate_model = RecordTableModel(
            [ColumnSpec(key, header, Qt.AlignLeft if key in {"part_number", "assigned_zones", "source_slot", "reason"} else Qt.AlignCenter, 170) for key, header in DUPLICATE_PLAN_COLUMNS],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.duplicate_model)
        self.duplicate_table = QTableView()
        configure_table(self.duplicate_table, self.duplicate_model, wrap_headers=True)
        install_copy_menu(self.duplicate_table, self.duplicate_model)
        page.layout.addWidget(self.duplicate_table, 1)
        return page

    def _build_warning_tab(self):
        page = Card()
        header = QLabel("Warnings / Validation")
        header.setObjectName("SectionTitle")
        page.layout.addWidget(header)
        self.warning_model = RecordTableModel(
            [ColumnSpec(key, header, Qt.AlignLeft, 180 if key == "severity" else 720) for key, header in WARNING_COLUMNS],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.warning_model)
        self.warning_table = QTableView()
        configure_table(self.warning_table, self.warning_model, wrap_headers=True)
        install_copy_menu(self.warning_table, self.warning_model)
        page.layout.addWidget(self.warning_table, 1)
        return page

    def browse_source(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Feeder Mapping Excel/CSV",
            "",
            "Feeder Data (*.xlsx *.xlsm *.xls *.csv);;Excel Workbook (*.xlsx *.xlsm *.xls);;CSV (*.csv);;All Files (*)",
        )
        if file_path:
            self.source_picker.set_path(file_path)
            self.load_preview()

    def load_preview(self):
        source_path = self.source_picker.path()
        if not source_path:
            return
        self._clear_result_tables(keep_preview=False)
        self.run_worker(
            lambda path=source_path: load_feeder_balancer_preview(path),
            self._on_preview_loaded,
            "Loading feeder file...",
        )

    def _on_preview_loaded(self, result):
        self.preview_result = result
        self.balance_result = None
        self.detected_columns_label.setText(f"{len(result.columns)} columns | {result.row_count} rows | {result.sheet_name}")
        self.status_label.setText(f"Loaded: {result.source_file}")
        self._populate_mapping_combos(result.columns, result.suggested_mapping)
        self._set_preview_rows(result)
        self.preview_btn.setEnabled(True)
        self.generate_btn.setEnabled(True)
        self.export_btn.setEnabled(False)
        self.summary_label.setText("0 ZONES | 0 PARTS")

    def _populate_mapping_combos(self, columns, suggested_mapping):
        for field_name, combo in self.mapping_combos.items():
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("Not used")
            combo.addItems(columns)
            selected = suggested_mapping.get(field_name, "")
            if selected:
                combo.setCurrentText(selected)
            elif field_name in {"slot", "part_number", "component_insert"}:
                combo.setCurrentIndex(0)
            combo.blockSignals(False)

    def _set_preview_rows(self, result):
        preview_columns = list(result.preview_rows[0].keys()) if result.preview_rows else result.columns[:12]
        self.preview_model.set_columns([ColumnSpec(column, column, Qt.AlignLeft, 150) for column in preview_columns])
        self.preview_model.set_records(result.preview_rows)

    def preview_detection(self):
        if not self._validate_source_loaded():
            return
        self.run_worker(
            lambda: detect_feeder_balancer_zones(self._config()),
            self._on_detection_done,
            "Detecting feeder zones...",
        )

    def _on_detection_done(self, result):
        self.zone_result = result
        self.zone_model.set_records(result.zone_records)
        self._set_warning_rows(result.warnings)
        self.summary_label.setText(f"{result.detected_zone_count} ZONES | {result.part_count} PARTS | {result.slot_count} SLOTS")
        self.status_label.setText("Detection ready")

    def generate_balance(self):
        if not self._validate_source_loaded():
            return
        self.balance_result = None
        self.export_btn.setEnabled(False)

        worker = None

        def task():
            return analyze_feeder_balance(
                self._config(),
                progress_callback=lambda percent, message: worker.signals.progress.emit(percent, message),
            )

        worker = TaskWorker(task)
        worker._busy_text = "Generating feeder balance..."
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, "Generating feeder balance..."))
        worker.signals.progress.connect(self._on_progress)
        worker.signals.result.connect(self._on_balance_done)
        worker.signals.error.connect(self._show_worker_error)
        worker.signals.finished.connect(lambda w=worker: self._finish_worker(w))
        self.thread_pool.start(worker)

    def _on_progress(self, percent, message):
        self.status_label.setText(message)
        self.progress.setVisible(True)
        self.progress.setRange(0, 100)
        self.progress.setValue(max(0, min(100, percent)))

    def _on_balance_done(self, result):
        self.balance_result = result
        self.balanced_model.set_records(result.balanced_rows)
        self.summary_model.set_records(result.zone_summary_rows)
        self.duplicate_model.set_records(result.duplicate_plan_rows)
        self._set_warning_rows(result.warnings)
        self.zone_model.set_records(self._zone_records_from_summary(result))
        self.export_btn.setEnabled(True)
        self.summary_label.setText(
            f"{len(result.zone_summary_rows)} ZONES | {result.metrics.get('unique_part_count', 0)} PARTS | "
            f"{result.metrics.get('duplicate_feeder_count', 0)} DUP"
        )
        self.status_label.setText(f"Done: {result.optimization_status}")

    def export_excel(self):
        if self.balance_result is None:
            QMessageBox.information(self, "Export Excel", "Belum ada hasil balancing untuk diexport.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Feeder Balancer Workbook",
            suggest_export_name(self.balance_result.source_path),
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        try:
            saved_path = export_feeder_balance_result(self.balance_result, output_path)
        except Exception as exc:
            QMessageBox.warning(self, "Export gagal", str(exc))
            return

        self.status_label.setText(f"Exported: {Path(saved_path).name}")
        self.status_label.setToolTip(saved_path)
        QMessageBox.information(self, "Export Excel", f"File berhasil dibuat:\n{saved_path}")

    def clear_data(self):
        self.preview_result = None
        self.zone_result = None
        self.balance_result = None
        self.source_picker.clear()
        self.detected_columns_label.setText("No file loaded")
        self.status_label.setText("Ready")
        self.status_label.setToolTip("")
        self.summary_label.setText("0 ZONES | 0 PARTS")
        self.profile_input.clear()
        self.duplicate_min_spin.setValue(int(DEFAULT_DUPLICATE_MIN_INSERT))
        self.multi_copy_spin.setValue(int(DEFAULT_MULTI_COPY_MIN_INSERT))
        for combo in self.mapping_combos.values():
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("Not used")
            combo.blockSignals(False)
        self._clear_result_tables(keep_preview=False)
        self.preview_btn.setEnabled(False)
        self.generate_btn.setEnabled(False)
        self.export_btn.setEnabled(False)

    def _clear_result_tables(self, keep_preview=True):
        if not keep_preview:
            self.preview_model.set_records([])
        self.zone_model.set_records([])
        self.balanced_model.set_records([])
        self.summary_model.set_records([])
        self.duplicate_model.set_records([])
        self.warning_model.set_records([])

    def _set_warning_rows(self, warnings):
        self.warning_model.set_records([{"severity": "WARNING", "message": warning} for warning in warnings])

    def _zone_records_from_summary(self, result):
        return [
            {
                "zone": row.get("zone", ""),
                "slot_count": row.get("slot_count", ""),
                "usable_slot_count": row.get("usable_slot_count", ""),
                "parsed_zone_count": "",
                "parsed_slot_number_count": "",
                "parsed_side_count": "",
                "parse_status": "BALANCED",
                "sample_slots": f"Assigned feeders: {row.get('assigned_feeder_count', '')}",
            }
            for row in result.zone_summary_rows
        ]

    def _validate_source_loaded(self):
        if not self.source_picker.path():
            QMessageBox.warning(self, "Input belum lengkap", "File Excel/CSV belum dipilih.")
            return False
        if self.preview_result is None:
            QMessageBox.warning(self, "Input belum diload", "Load file terlebih dahulu lewat tombol Browse.")
            return False
        return True

    def _config(self):
        return FeederBalancerConfig(
            source_path=self.source_picker.path(),
            machine_mode=self.machine_combo.currentText(),
            column_mapping=self._current_mapping(),
            profile_text=self.profile_input.toPlainText() if self.machine_combo.currentText() == MACHINE_NPM_CUSTOM else "",
            duplicate_min_insert=self.duplicate_min_spin.value(),
            multi_copy_min_insert=self.multi_copy_spin.value(),
        )

    def _current_mapping(self):
        mapping = {}
        for field_name, combo in self.mapping_combos.items():
            value = combo.currentText()
            mapping[field_name] = "" if value == "Not used" else value
        return mapping

    def _sync_profile_ui(self):
        visible = self.machine_combo.currentText() == MACHINE_NPM_CUSTOM
        self.profile_label.setVisible(visible)
        self.profile_input.setVisible(visible)
        self._on_mapping_changed()

    def _on_mapping_changed(self, *_):
        self.balance_result = None
        self.export_btn.setEnabled(False)
        if self.preview_result is not None:
            self.zone_model.set_records([])
            self.summary_model.set_records([])
            self.duplicate_model.set_records([])
            self.warning_model.set_records([])
            self.status_label.setText("Mapping changed")
