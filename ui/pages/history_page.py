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
from services import history_service
from ui.pages.base import WorkerPage
from widgets.card import Card
from widgets.table_tools import configure_table, install_copy_menu


class HistoryPage(WorkerPage):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(thread_pool, theme_manager, parent)
        self.history_data = []
        self._build_ui()
        self.refresh_history()
        self.theme_manager.changed.connect(self.apply_theme_to_models)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        header = QHBoxLayout()
        self.refresh_btn = QPushButton("Refresh List")
        self.refresh_btn.setObjectName("PrimaryButton")
        self.refresh_btn.clicked.connect(self.refresh_history)
        self.clear_btn = QPushButton("Clear All History")
        self.clear_btn.setObjectName("DangerButton")
        self.clear_btn.clicked.connect(self.clear_history)
        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedLabel")
        header.addWidget(self.refresh_btn)
        header.addStretch(1)
        header.addWidget(self.status_label)
        header.addWidget(self.clear_btn)
        root.addLayout(header)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        card = Card()
        self.history_model = RecordTableModel(
            [
                ColumnSpec("timestamp", "Timestamp", Qt.AlignCenter, 170),
                ColumnSpec("txt_file", "Reference / Machine", Qt.AlignLeft, 260),
                ColumnSpec("tsv_file", "Source / Program", Qt.AlignLeft, 260),
                ColumnSpec("stats", "Changes (A/C/D)", Qt.AlignCenter, 170),
            ],
            theme=self.theme_manager.theme,
        )
        self.register_model(self.history_model)
        self.table = QTableView()
        configure_table(self.table, self.history_model)
        install_copy_menu(self.table, self.history_model)
        card.layout.addWidget(self.table, 1)
        root.addWidget(card, 1)

        self.detail_btn = QPushButton("View Details / Export Selected")
        self.detail_btn.setObjectName("PrimaryButton")
        self.detail_btn.clicked.connect(self.export_selected)
        root.addWidget(self.detail_btn)

        self.register_busy_widgets(self.refresh_btn, self.clear_btn, self.detail_btn)

    def refresh_history(self):
        self.history_data = history_service.load_history()
        records = []
        for entry in self.history_data:
            records.append(
                {
                    "timestamp": entry.get("timestamp", ""),
                    "txt_file": entry.get("txt_file", ""),
                    "tsv_file": entry.get("tsv_file", ""),
                    "stats": f"{entry.get('add_count', 0)} ADD | {entry.get('cng_count', 0)} CNG | {entry.get('del_count', 0)} DEL",
                }
            )
        self.history_model.set_records(records)
        self.status_label.setText(f"{len(records)} history item(s)")

    def clear_history(self):
        if QMessageBox.question(self, "Confirm", "Are you sure you want to clear all history?") != QMessageBox.Yes:
            return
        try:
            history_service.clear_history()
            self.refresh_history()
        except Exception as exc:
            QMessageBox.warning(self, "Error", str(exc))

    def export_selected(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.warning(self, "Warning", "Select a history entry first!")
            return

        row = rows[0].row()
        if row < 0 or row >= len(self.history_data):
            QMessageBox.warning(self, "Warning", "Selected history entry is invalid.")
            return

        entry = self.history_data[row]
        if QMessageBox.question(self, "Export", f"Do you want to export results from {entry.get('timestamp', '')} to Excel?") != QMessageBox.Yes:
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "Export History", "", "Excel (*.xlsx)")
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        self.run_worker(
            lambda item=entry, path=file_path: history_service.export_history_entry(item, path),
            lambda _, path=file_path: QMessageBox.information(self, "Success", f"Exported to:\n{path}"),
            "Exporting history...",
        )
