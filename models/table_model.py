from dataclasses import dataclass

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QColor


@dataclass(frozen=True)
class ColumnSpec:
    key: str
    header: str
    alignment: Qt.AlignmentFlag = Qt.AlignCenter
    width: int = 120


class RecordTableModel(QAbstractTableModel):
    def __init__(self, columns=None, records=None, status_key=None, theme=None, parent=None):
        super().__init__(parent)
        self.columns = columns or []
        self.records = records or []
        self.status_key = status_key
        self.theme = theme or {}

    def rowCount(self, parent=QModelIndex()):
        if parent.isValid():
            return 0
        return len(self.records)

    def columnCount(self, parent=QModelIndex()):
        if parent.isValid():
            return 0
        return len(self.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        row = self.records[index.row()]
        column = self.columns[index.column()]
        value = row.get(column.key, "")

        if role == Qt.DisplayRole:
            return "" if value is None else str(value)
        if role == Qt.TextAlignmentRole:
            return int(column.alignment | Qt.AlignVCenter)
        if role == Qt.BackgroundRole:
            status = self._status(row)
            key = {
                "ADD": "add_bg",
                "CNG": "cng_bg",
                "DEL": "del_bg",
                "MATCH": "match_bg",
                "SAFE": "match_bg",
                "CONFLICT": "del_bg",
                "CHECK": "cng_bg",
                "GROUPED": "match_bg",
                "SINGLE": "cng_bg",
            }.get(status)
            return QColor(self.theme.get(key)) if key and self.theme.get(key) else None
        if role == Qt.ForegroundRole:
            diff_keys = row.get("_diff_keys", [])
            if column.key in diff_keys and self.theme.get("red"):
                return QColor(self.theme["red"])
            status = self._status(row)
            key = {
                "ADD": "add_fg",
                "CNG": "cng_fg",
                "DEL": "del_fg",
                "MATCH": "match_fg",
                "SAFE": "match_fg",
                "CONFLICT": "del_fg",
                "CHECK": "cng_fg",
                "GROUPED": "match_fg",
                "SINGLE": "cng_fg",
            }.get(status)
            return QColor(self.theme.get(key)) if key and self.theme.get(key) else None
        if role == Qt.UserRole:
            return value
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal and 0 <= section < len(self.columns):
            return self.columns[section].header
        return section + 1

    def set_records(self, records):
        self.beginResetModel()
        self.records = records or []
        self.endResetModel()

    def set_columns(self, columns):
        self.beginResetModel()
        self.columns = columns or []
        self.endResetModel()

    def set_theme(self, theme):
        self.theme = theme or {}
        if self.records:
            top_left = self.index(0, 0)
            bottom_right = self.index(len(self.records) - 1, len(self.columns) - 1)
            self.dataChanged.emit(top_left, bottom_right, [Qt.BackgroundRole, Qt.ForegroundRole])

    def selected_rows_as_tsv(self, row_numbers, include_headers=True, clean_copy=False, excluded_keys=None):
        rows = [self.records[i] for i in row_numbers if 0 <= i < len(self.records)]
        return self.records_as_tsv(rows, include_headers=include_headers, clean_copy=clean_copy, excluded_keys=excluded_keys)

    def all_as_tsv(self, include_headers=True, clean_copy=False, excluded_keys=None):
        return self.records_as_tsv(self.records, include_headers=include_headers, clean_copy=clean_copy, excluded_keys=excluded_keys)

    def records_as_tsv(self, rows, include_headers=True, clean_copy=False, excluded_keys=None):
        columns = self.columns
        if clean_copy and len(columns) > 2:
            columns = columns[1:-1]
        if excluded_keys:
            excluded_keys = set(excluded_keys)
            columns = [column for column in columns if column.key not in excluded_keys]

        output = []
        if include_headers and not clean_copy:
            output.append("\t".join(col.header for col in columns))
        for row in rows:
            output.append("\t".join(str(row.get(col.key, "")) for col in columns))
        return "\n".join(output)

    def _status(self, row):
        if not self.status_key:
            return None
        status = str(row.get(self.status_key, "")).upper()
        if status == "MOD":
            return "CNG"
        if "BEDA" in status:
            return "CNG"
        if "ADD" in status:
            return "ADD"
        if "REMOVE" in status:
            return "DEL"
        return status
