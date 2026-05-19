from PySide6.QtGui import QColor, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (
    QFileDialog,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
)

from services import worksheet_service
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.file_picker import FilePicker


FILTER_NG_ID = "NG - BEDA ID / P/N (Default)"
FILTER_NG_QTY = "NG - BEDA JUMLAH / QTY"
FILTER_ALL_NG = "Tampilkan Semua NG"
FILTER_MATCH = "Tampilkan MATCH (OK)"


class WorksheetComparatorPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.ng_id = []
        self.ng_qty = []
        self.matches = []
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Worksheet vs CRB Verification")
        title.setObjectName("SectionTitle")
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.status_label)
        root.addLayout(header)

        input_card = Card()
        self.excel_picker = FilePicker("WORKSHEET (Excel/CSV):")
        self.excel_picker.browse_requested.connect(self.browse_excel)
        self.crb_picker = FilePicker("NPM (.crb):")
        self.crb_picker.browse_requested.connect(self.browse_crb)
        input_card.layout.addWidget(self.excel_picker)
        input_card.layout.addWidget(self.crb_picker)
        root.addWidget(input_card)

        action_bar = QHBoxLayout()
        self.compare_btn = QPushButton("Compare Data")
        self.compare_btn.setObjectName("SuccessButton")
        self.compare_btn.clicked.connect(self.run_compare)
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_data)
        action_bar.addWidget(self.compare_btn)
        action_bar.addWidget(self.clear_btn)
        action_bar.addStretch(1)
        root.addLayout(action_bar)

        filter_bar = QHBoxLayout()
        filter_label = QLabel("Filter Hasil:")
        filter_label.setObjectName("SectionTitle")
        self.status_filter = QComboBox()
        self.status_filter.addItems([FILTER_NG_ID, FILTER_NG_QTY, FILTER_ALL_NG, FILTER_MATCH])
        self.status_filter.currentTextChanged.connect(self.display_results)
        filter_bar.addWidget(filter_label)
        filter_bar.addWidget(self.status_filter)
        filter_bar.addStretch(1)
        root.addLayout(filter_bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        root.addWidget(self.result_text, 1)

        self.register_busy_widgets(self.compare_btn, self.clear_btn, self.excel_picker.button, self.crb_picker.button)

    def browse_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Pilih File Excel", "", "Excel Files (*.xlsx *.xls *.csv)")
        if file_path:
            self.excel_picker.set_path(file_path)

    def browse_crb(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Pilih File CRB", "", "CRB Files (*.crb *.CRB);;All Files (*)")
        if file_path:
            self.crb_picker.set_path(file_path)

    def run_compare(self):
        excel_path = self.excel_picker.path()
        crb_path = self.crb_picker.path()
        self.result_text.clear()
        self.append_text("=== MENGOLAH DATA EXCEL ===")
        self.append_text("=== MENGOLAH DATA MESIN (CRB) ===")
        self.run_worker(lambda: worksheet_service.run_worksheet_compare(excel_path, crb_path), self._on_compare_done, "Running worksheet compare...")

    def _on_compare_done(self, result):
        self.ng_id = result.ng_id
        self.ng_qty = result.ng_qty
        self.matches = result.matches
        total_ng = len(self.ng_id) + len(self.ng_qty)

        if total_ng > 0:
            self.status_filter.setCurrentText(FILTER_NG_ID)
            self.display_results()
            QMessageBox.warning(self, "Compare Selesai", f"Ditemukan {total_ng} Data NG!\nSilakan cek log layar.")
        else:
            self.result_text.clear()
            self.append_text("")
            self.append_text("SEMUA DATA MATCH! (0 NG)")
            QMessageBox.information(self, "Hasil Compare", "OK COMPARE")
        self.status_label.setText("Done")

    def clear_data(self):
        self.excel_picker.clear()
        self.crb_picker.clear()
        self.result_text.clear()
        self.ng_id.clear()
        self.ng_qty.clear()
        self.matches.clear()
        self.status_filter.setCurrentText(FILTER_NG_ID)
        self.append_text("=== KOLOM BERHASIL DIBERSIHKAN. SILAKAN MASUKKAN FILE BARU ===")
        self.status_label.setText("")

    def display_results(self, filter_value=None):
        filter_value = filter_value or self.status_filter.currentText()
        self.result_text.clear()

        if not self.ng_id and not self.ng_qty and not self.matches:
            return

        if not self.ng_id and not self.ng_qty:
            self.append_text("")
            self.append_text("SEMUA DATA MATCH! (0 NG)")
            return

        self.append_text(f"=== HASIL COMPARE (FILTER: {filter_value}) ===")
        self.append_text("")

        if filter_value == FILTER_NG_ID:
            if self.ng_id:
                self.append_text(">>> DAFTAR BEDA / ERROR ID / P/N (NG) <<<")
                for line in self.ng_id:
                    self.append_text(line)
            else:
                self.append_text("Tidak ada NG untuk Beda ID / P/N.")
        elif filter_value == FILTER_NG_QTY:
            if self.ng_qty:
                self.append_text(">>> DAFTAR BEDA JUMLAH / QTY (NG) <<<")
                for line in self.ng_qty:
                    self.append_text(line)
            else:
                self.append_text("Tidak ada NG untuk Beda Jumlah / QTY.")
        elif filter_value == FILTER_ALL_NG:
            if self.ng_id:
                self.append_text(">>> DAFTAR BEDA / ERROR ID / P/N (NG) <<<")
                for line in self.ng_id:
                    self.append_text(line)
                self.append_text("")
                self.append_text("--------------------------------------------------")
                self.append_text("")
            if self.ng_qty:
                self.append_text(">>> DAFTAR BEDA JUMLAH / QTY (NG) <<<")
                for line in self.ng_qty:
                    self.append_text(line)
        elif filter_value == FILTER_MATCH:
            if self.matches:
                self.append_text(">>> DAFTAR COCOK (MATCH) <<<")
                for line in self.matches:
                    self.append_text(line)
            else:
                self.append_text("Tidak ada data yang MATCH.")

    def append_text(self, text):
        cursor = self.result_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        fmt = QTextCharFormat()
        theme = self.theme_manager.theme

        if "[NG]" in text or "ERROR" in text or "DAFTAR BEDA" in text:
            fmt.setForeground(QColor(theme["del_fg"]))
            fmt.setFontWeight(700)
        elif "[MATCH]" in text or "AMAN" in text or "0 NG" in text or "SEMUA DATA MATCH" in text:
            fmt.setForeground(QColor(theme["add_fg"]))
        elif "===" in text or ">>>" in text or "---" in text:
            fmt.setForeground(QColor(theme["orange"]))
            fmt.setFontWeight(700)
        else:
            fmt.setForeground(QColor(theme["text"]))

        cursor.insertText(text + "\n", fmt)
        self.result_text.setTextCursor(cursor)
        self.result_text.ensureCursorVisible()
