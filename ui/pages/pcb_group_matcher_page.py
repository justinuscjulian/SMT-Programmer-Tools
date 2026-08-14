from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QButtonGroup,
    QDialog,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSplitter,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from models.table_model import ColumnSpec, RecordTableModel
from services.pcb_group_matcher_service import (
    MODE_BOM_FILE,
    MODE_PROGRAM_EXCEL,
    MODE_PROGRAM_FOLDER,
    PcbGroupMatcherConfig,
    analyze_pcb_group_matcher,
    export_pcb_group_matcher_result,
    suggest_output_name,
)
from ui.pages.base import WorkerPage, TaskWorker
from services.pcb_group_matcher_service import generate_merged_fix_feeder_group
from widgets.card import Card
from widgets.file_picker import FilePicker
from widgets.table_tools import configure_table, install_copy_menu


class PcbGroupMatcherPage(WorkerPage):
    progress_updated = Signal(int, str)

    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.matcher_result = None
        self._build_ui()
        self.theme_manager.changed.connect(self.apply_theme_to_models)
        self.progress_updated.connect(self._on_progress_update)

    def _on_progress_update(self, percent, msg):
        self.progress.setValue(percent)
        self.status_label.setText(msg)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        # Header
        header = QHBoxLayout()
        title = QLabel("PCB Group Matcher")
        title.setObjectName("SectionTitle")
        self.summary_label = QLabel("0 GROUPS MATCHED")
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

        # Source Card
        source_card = Card()
        source_title = QLabel("PCB Input & Fix Feeder Source")
        source_title.setObjectName("SectionTitle")
        source_card.layout.addWidget(source_title)

        # Import Mode Radios
        mode_row = QHBoxLayout()
        mode_label = QLabel("Pilih Metode Import PCB Baru:")
        mode_label.setStyleSheet("font-weight: bold;")
        mode_row.addWidget(mode_label)

        self.mode_bom_radio = QRadioButton("Import Single BOM File")
        self.mode_folder_radio = QRadioButton("Import PCB Folder")
        self.mode_excel_radio = QRadioButton("Import Single Program Excel")
        self.mode_bom_radio.setChecked(True)

        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.mode_bom_radio)
        self.mode_group.addButton(self.mode_folder_radio)
        self.mode_group.addButton(self.mode_excel_radio)

        self.mode_bom_radio.toggled.connect(self._sync_mode_ui)
        self.mode_folder_radio.toggled.connect(self._sync_mode_ui)
        self.mode_excel_radio.toggled.connect(self._sync_mode_ui)

        mode_row.addWidget(self.mode_bom_radio)
        mode_row.addWidget(self.mode_folder_radio)
        mode_row.addWidget(self.mode_excel_radio)
        mode_row.addStretch(1)
        source_card.layout.addLayout(mode_row)

        # PCB File / Folder Picker
        self.pcb_picker = FilePicker("File BOM PCB Baru (.xlsx, .xls, .csv, .tsv, .txt):")
        self.pcb_picker.browse_requested.connect(self.browse_pcb_source)
        source_card.layout.addWidget(self.pcb_picker)

        # Fix Feeder Group Excel Picker
        self.group_picker = FilePicker("File Fix Feeder Group Excel (Output dari All Table Fix Feeder):")
        self.group_picker.browse_requested.connect(self.browse_group_file)
        source_card.layout.addWidget(self.group_picker)

        line_type_layout = QHBoxLayout()
        line_type_label = QLabel("Tipe Line Mesin (Penting untuk mode Program Folder/Excel):")
        line_type_label.setStyleSheet("font-weight: bold;")
        self.line_type_combo = QComboBox()
        self.line_type_combo.addItems(["Line 1-5", "Line 6-7", "Line 8", "Line 9", "CM602"])
        line_type_layout.addWidget(line_type_label)
        line_type_layout.addWidget(self.line_type_combo)
        line_type_layout.addStretch()
        source_card.layout.addLayout(line_type_layout)

        root.addWidget(source_card)

        # Action Bar
        action_bar = QHBoxLayout()
        self.analyze_btn = QPushButton("Analyze PCB Group Match")
        self.analyze_btn.setObjectName("PrimaryButton")
        self.analyze_btn.clicked.connect(self.run_analysis)

        self.export_btn = QPushButton("Export Result Excel")
        self.export_btn.setObjectName("SuccessButton")
        self.export_btn.setEnabled(False)

        self.merge_btn = QPushButton("Merge & Generate Fix Feeder")
        self.merge_btn.setObjectName("PrimaryButton")
        self.merge_btn.setEnabled(False)
        self.merge_btn.clicked.connect(self.show_merge_dialog)

        self.export_btn.clicked.connect(self.export_result)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_data)

        action_bar.addWidget(self.analyze_btn)
        action_bar.addWidget(self.export_btn)
        action_bar.addWidget(self.merge_btn)
        action_bar.addWidget(self.clear_btn)
        action_bar.addStretch(1)

        root.addLayout(action_bar)

        # Table & Detail Splitter
        splitter = QSplitter(Qt.Vertical)

        # Ranking Table Card
        table_card = Card()
        table_title = QLabel("Ranking Group Match Results")
        table_title.setObjectName("SectionTitle")
        table_card.layout.addWidget(table_title)

        self.result_model = RecordTableModel(
            [
                ColumnSpec("rank", "Rank", Qt.AlignCenter, 70),
                ColumnSpec("group_name", "Group Name", Qt.AlignLeft, 140),
                ColumnSpec("match_rate_text", "Match Rate %", Qt.AlignCenter, 130),
                ColumnSpec("matched_count_text", "Matched Parts", Qt.AlignCenter, 140),
                ColumnSpec("missing_count", "Extra Needed", Qt.AlignCenter, 130),
                ColumnSpec("member_pcbs_text", "Member PCBs in Group", Qt.AlignLeft, 240),
                ColumnSpec("status_recommendation", "Recommendation", Qt.AlignLeft, 280),
            ],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.result_model)
        self.result_table = QTableView()
        configure_table(self.result_table, self.result_model, wrap_headers=True)
        install_copy_menu(self.result_table, self.result_model)
        self.result_table.selectionModel().selectionChanged.connect(self._on_table_selection_changed)

        table_card.layout.addWidget(self.result_table, 1)
        splitter.addWidget(table_card)

        # Details Display Card
        detail_card = Card()
        detail_title = QLabel("Detail Komponen Matched & Extra Feeder Needed")
        detail_title.setObjectName("SectionTitle")
        detail_card.layout.addWidget(detail_title)

        self.detail_text = QTextEdit()
        self.detail_text.setReadOnly(True)
        self.detail_text.setPlaceholderText("Pilih baris Group pada tabel di atas untuk melihat rincian komponen yang ter-cover dan feeder tambahan yang dibutuhkan...")
        detail_card.layout.addWidget(self.detail_text, 1)
        splitter.addWidget(detail_card)

        root.addWidget(splitter, 1)

        self.register_busy_widgets(
            self.analyze_btn,
            self.export_btn,
            self.merge_btn,
            self.clear_btn,
            self.pcb_picker.button,
            self.group_picker.button,
            self.mode_bom_radio,
            self.mode_folder_radio,
            self.mode_excel_radio,
        )

        self._sync_mode_ui()

    def _get_selected_mode(self):
        if self.mode_folder_radio.isChecked():
            return MODE_PROGRAM_FOLDER
        if self.mode_excel_radio.isChecked():
            return MODE_PROGRAM_EXCEL
        return MODE_BOM_FILE

    def _sync_mode_ui(self):
        mode = self._get_selected_mode()
        if mode == MODE_BOM_FILE:
            self.pcb_picker.label.setText("File BOM PCB Baru (.xlsx, .xls, .csv, .tsv, .txt):")
            self.pcb_picker.button.setText("Browse File")
        elif mode == MODE_PROGRAM_FOLDER:
            self.pcb_picker.label.setText("Folder Induk Program PCB Baru:")
            self.pcb_picker.button.setText("Browse Folder")
        elif mode == MODE_PROGRAM_EXCEL:
            self.pcb_picker.label.setText("File Program Excel PCB Baru (.xlsx, .xls, .xlsm):")
            self.pcb_picker.button.setText("Browse File")

        self.pcb_picker.clear()

    def browse_pcb_source(self):
        mode = self._get_selected_mode()
        if mode == MODE_PROGRAM_FOLDER:
            folder_path = QFileDialog.getExistingDirectory(
                self,
                "Pilih Folder Induk Program PCB Baru",
                "",
            )
            if folder_path:
                self.pcb_picker.set_path(folder_path)
        elif mode == MODE_PROGRAM_EXCEL:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Pilih File Program Excel PCB Baru",
                "",
                "Excel Files (*.xlsx *.xls *.xlsm *.xlsb);;All Files (*)",
            )
            if file_path:
                self.pcb_picker.set_path(file_path)
        else:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Pilih File BOM PCB Baru",
                "",
                "BOM Files (*.xlsx *.xls *.csv *.tsv *.txt);;Excel Files (*.xlsx *.xls);;Text Files (*.csv *.tsv *.txt);;All Files (*)",
            )
            if file_path:
                self.pcb_picker.set_path(file_path)

    def browse_group_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Pilih File Fix Feeder Group Excel",
            "",
            "Excel Files (*.xlsx *.xlsm);;All Files (*)",
        )
        if file_path:
            self.group_picker.set_path(file_path)

    def run_analysis(self):
        pcb_path = self.pcb_picker.path()
        group_path = self.group_picker.path()

        if not pcb_path:
            QMessageBox.warning(self, "Input belum lengkap", "Source PCB Baru belum dipilih.")
            return
        if not group_path:
            QMessageBox.warning(self, "Input belum lengkap", "File Fix Feeder Group Excel belum dipilih.")
            return

        config = PcbGroupMatcherConfig(
            import_mode=self._get_selected_mode(),
            pcb_source_path=pcb_path,
            fix_feeder_group_path=group_path,
            line_type=self.line_type_combo.currentText(),
        )

        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.status_label.setText("Analyzing PCB Group Match...")

        def _worker_progress(percent, msg):
            self.progress_updated.emit(percent, msg)

        self.run_worker(
            lambda: analyze_pcb_group_matcher(config, progress_callback=_worker_progress),
            self._on_analysis_done,
            "Analyzing PCB Group Match...",
        )

    def _on_analysis_done(self, result):
        self.progress.setVisible(False)
        self.matcher_result = result

        table_records = []
        for item in result.group_matches:
            table_records.append({
                "rank": item.rank,
                "group_name": item.group_name,
                "match_rate_text": f"{item.match_rate_percent}%",
                "matched_count_text": f"{item.matched_count} / {item.total_pcb_parts}",
                "missing_count": item.missing_count,
                "member_pcbs_text": ", ".join(item.member_pcbs) if item.member_pcbs else "-",
                "status_recommendation": item.status_recommendation,
                "_raw_item": item,
            })

        self.result_model.set_records(table_records)
        self.summary_label.setText(f"{len(result.group_matches)} GROUPS MATCHED | PCB: {result.pcb_parts_count} PARTS")
        self.export_btn.setEnabled(True)
        if self.result_table.model().rowCount() > 0:
            self.result_table.selectRow(0)

        if result.best_match:
            self.status_label.setText(f"Done: {result.best_match.group_name} Best Match ({result.best_match.match_rate_percent}%)")
            QMessageBox.information(
                self,
                "PCB Group Matcher Selesai",
                (
                    f"Hasil Analisis Kecocokan PCB [{result.pcb_name}]:\n\n"
                    f"🏆 Rekomendasi Group Terbaik: {result.best_match.group_name}\n"
                    f"📊 Match Rate: {result.best_match.match_rate_percent}%\n"
                    f"✅ Komponen Ter-cover: {result.best_match.matched_count} / {result.best_match.total_pcb_parts}\n"
                    f"🔧 Feeder Tambahan Dibutuhkan: {result.best_match.missing_count}\n\n"
                    f"Catatan:\n{result.best_match.recommendation_note}"
                ),
            )
        else:
            self.status_label.setText("Done: No match found")
            QMessageBox.warning(
                self,
                "PCB Group Matcher Selesai",
                f"Hasil Analisis Kecocokan PCB [{result.pcb_name}]:\n\nTidak ada group yang valid untuk dicocokkan."
            )

    def _on_table_selection_changed(self):
        indexes = self.result_table.selectionModel().selectedRows()
        if not indexes or self.matcher_result is None:
            self.detail_text.clear()
            return

        row_idx = indexes[0].row()
        records = self.result_model.records
        if 0 <= row_idx < len(records):
            item = records[row_idx].get("_raw_item")
            if item:
                self._display_group_detail(item)

    def _display_group_detail(self, item):
        lines = []
        lines.append(f"=== DETIL KECOCOKAN: {item.group_name} (Rank #{item.rank}) ===")
        lines.append(f"• Member PCBs di Group : {', '.join(item.member_pcbs) if item.member_pcbs else '-'}")
        lines.append(f"• Kecocokan (Match Rate): {item.match_rate_percent}% ({item.matched_count} dari {item.total_pcb_parts} komponen)")
        lines.append(f"• Feeder Ekstra Dibutuhkan: {item.missing_count} komponen")
        lines.append(f"• Status & Catatan : {item.status_recommendation}")
        lines.append(f"  {item.recommendation_note}\n")

        lines.append(f"--- 1. DAFTAR KOMPONEN YANG SUDAH TER-COVER DI FIX FEEDER ({len(item.matched_details)}) ---")
        if item.matched_details:
            for m in item.matched_details:
                lines.append(f"  [Slot {m['location_code']}] {m['part_number']} ({m['type']})")
        else:
            lines.append("  (Tidak ada komponen yang ter-cover)")

        lines.append(f"\n--- 2. DAFTAR FEEDER TAMBAHAN YANG HARUS DIPASANG MANUAL / CHANGE-OVER ({len(item.missing_details)}) ---")
        if item.missing_details:
            for p in item.missing_details:
                lines.append(f"  • {p}  --> (Pasang sebagai Feeder Ekstra di Slot Kosong)")
        else:
            lines.append("  (Semua komponen PCB baru sudah ter-cover sempurna!)")

        self.detail_text.setText("\n".join(lines))

    def export_result(self):
        if not self.matcher_result:
            QMessageBox.warning(self, "Export Belum Siap", "Belum ada hasil analisis kecocokan.")
            return

        suggested = suggest_output_name(self.matcher_result.pcb_name)
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save PCB Group Matcher Excel Report",
            suggested,
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        self.run_worker(
            lambda: export_pcb_group_matcher_result(self.matcher_result, output_path),
            self._on_export_done,
            "Exporting PCB Group Matcher report...",
        )

    def _on_export_done(self, output_file):
        output_name = Path(output_file).name
        self.status_label.setText(f"Exported: {output_name}")
        self.status_label.setToolTip(output_file)

        QMessageBox.information(
            self,
            "Export Selesai",
            f"Laporan analisis PCB Group Matcher berhasil disimpan di:\n{output_file}",
        )

    def clear_data(self):
        self.matcher_result = None
        self.pcb_picker.clear()
        self.group_picker.clear()
        self.result_model.set_records([])
        self.detail_text.clear()
        self.summary_label.setText("0 GROUPS MATCHED")
        self.status_label.setText("Ready")
        self.export_btn.setEnabled(False)
        self.merge_btn.setEnabled(False)

    def show_merge_dialog(self):
        indexes = self.result_table.selectionModel().selectedRows()
        if not indexes or not self.matcher_result:
            return
            
        match_item = self.matcher_result.group_matches[indexes[0].row()]
        
        dialog = MergeFeederDialog(match_item.group_name, self)
        dialog.line_type_combo.setCurrentText(self.line_type_combo.currentText())
        if dialog.exec() == QDialog.Accepted:
            line_type = dialog.line_type_combo.currentText()
            master_excel = dialog.master_picker.path()
            base_npm = dialog.base_picker.path()
            
            output_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Merged Fix Feeder Group",
                f"Merged_{match_item.group_name}.xlsx",
                "Excel Workbook (*.xlsx)",
            )
            if not output_path:
                return
                
            self._run_merge_task(match_item.group_name, line_type, master_excel, base_npm, output_path)

    def _run_merge_task(self, group_name, line_type, master_excel, base_npm, output_path):
        worker = None
        def task():
            return generate_merged_fix_feeder_group(
                self.matcher_result,
                group_name,
                line_type,
                master_excel,
                output_path,
                base_npm_path=base_npm,
                progress_callback=worker.signals.progress.emit
            )
            
        worker = TaskWorker(task)
        worker._busy_text = "Merging Fix Feeder Group..."
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, "Merging Fix Feeder Group..."))
        worker.signals.progress.connect(self._on_progress_update)
        worker.signals.result.connect(self._on_merge_done)
        worker.signals.error.connect(self._show_worker_error)
        worker.signals.finished.connect(lambda w=worker: self._finish_worker(w))
        self.thread_pool.start(worker)

    def _on_merge_done(self, result):
        saved_path, unassigned_count = result
        self.status_label.setText("Merge selesai!")
        if unassigned_count > 0:
            QMessageBox.warning(self, "Merge Selesai dengan Peringatan", f"Berhasil membuat grup merged.\nNamun ada {unassigned_count} komponen yang gagal ditempatkan karena kapasitas mesin tidak cukup atau tidak ada di Master Mapping.\nFile tersimpan di: {saved_path}")
        else:
            QMessageBox.information(self, "Merge Selesai", f"Berhasil membuat grup merged!\nSemua komponen sukses ditempatkan.\nFile tersimpan di: {saved_path}")


class MergeFeederDialog(QDialog):
    def __init__(self, group_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Merge to {group_name}")
        self.setMinimumWidth(500)
        
        layout = QVBoxLayout(self)
        
        line_layout = QHBoxLayout()
        line_layout.addWidget(QLabel("Tipe Line:"))
        self.line_type_combo = QComboBox()
        self.line_type_combo.addItems(["Line 1-5", "Line 6-7", "Line 8", "Line 9", "CM602"])
        line_layout.addWidget(self.line_type_combo)
        line_layout.addStretch()
        layout.addLayout(line_layout)
        
        self.master_picker = FilePicker("Master Mapping Excel:")
        self.master_picker.browse_requested.connect(self.browse_master)
        layout.addWidget(self.master_picker)
        
        self.base_picker = FilePicker("Base NPM File (.txt/.crb) (Opsional):")
        self.base_picker.browse_requested.connect(self.browse_base)
        layout.addWidget(self.base_picker)
        
        btn_layout = QHBoxLayout()
        self.ok_btn = QPushButton("Generate")
        self.ok_btn.setObjectName("PrimaryButton")
        self.ok_btn.clicked.connect(self.validate_and_accept)
        
        self.cancel_btn = QPushButton("Batal")
        self.cancel_btn.clicked.connect(self.reject)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addWidget(self.ok_btn)
        
        layout.addLayout(btn_layout)

    def browse_master(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Pilih File Master Mapping",
            "",
            "Excel Files (*.xlsx *.xls *.xlsm);;All Files (*)"
        )
        if path:
            self.master_picker.set_path(path)
            
    def browse_base(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Pilih Base NPM File",
            "",
            "Text Files (*.txt);;CRB Files (*.crb);;All Files (*)"
        )
        if path:
            self.base_picker.set_path(path)

    def validate_and_accept(self):
        if not self.master_picker.path():
            QMessageBox.warning(self, "Error", "Master Mapping Excel wajib diisi!")
            return
        self.accept()
