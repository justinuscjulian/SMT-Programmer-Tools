from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import QApplication, QAbstractItemView, QHeaderView, QMenu, QStyle, QStyleOptionHeader, QTableView


class WrappedHeaderView(QHeaderView):
    def __init__(self, orientation, parent=None, min_height=48):
        super().__init__(orientation, parent)
        self._min_height = min_height
        self.setDefaultAlignment(Qt.AlignCenter)

    def sectionSizeFromContents(self, logical_index):
        size = super().sectionSizeFromContents(logical_index)
        if self.orientation() == Qt.Horizontal:
            size.setHeight(max(size.height(), self._min_height))
        return size

    def paintSection(self, painter, rect, logical_index):
        if not rect.isValid():
            return

        option = QStyleOptionHeader()
        self.initStyleOption(option)
        option.rect = rect
        option.section = logical_index
        option.text = ""

        self.style().drawControl(QStyle.CE_HeaderSection, option, painter, self)

        text = self.model().headerData(logical_index, self.orientation(), Qt.DisplayRole)
        if text is None:
            return

        painter.save()
        painter.setFont(self.font())
        painter.setPen(option.palette.color(QPalette.ButtonText))
        painter.drawText(rect.adjusted(6, 4, -6, -4), Qt.AlignCenter | Qt.TextWordWrap, str(text))
        painter.restore()


def configure_table(table: QTableView, model, wrap_headers=False):
    if wrap_headers:
        table.setHorizontalHeader(WrappedHeaderView(Qt.Horizontal, table))

    table.setModel(model)
    table.setAlternatingRowColors(True)
    table.setSortingEnabled(False)
    table.setSelectionBehavior(QAbstractItemView.SelectRows)
    table.setSelectionMode(QAbstractItemView.ExtendedSelection)
    table.setEditTriggers(QAbstractItemView.NoEditTriggers)
    table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
    table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
    table.setWordWrap(False)
    table.setCornerButtonEnabled(False)
    table.verticalHeader().setVisible(False)
    table.verticalHeader().setDefaultSectionSize(34)
    table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
    table.horizontalHeader().setHighlightSections(False)
    table.verticalHeader().setHighlightSections(False)
    table.horizontalHeader().setStretchLastSection(True)
    if wrap_headers:
        table.horizontalHeader().setMinimumHeight(48)
    table.setShowGrid(False)

    for idx, column in enumerate(getattr(model, "columns", [])):
        table.setColumnWidth(idx, column.width)


def install_copy_menu(table, model, clean_copy=False, allow_cell_column=False, copy_all_excluded_keys=None):
    table.setContextMenuPolicy(Qt.CustomContextMenu)

    def selected_row_numbers():
        selection = table.selectionModel()
        if selection is None:
            return []
        rows = [index.row() for index in selection.selectedRows()]
        if rows:
            return sorted(set(rows))
        return sorted(set(index.row() for index in selection.selectedIndexes()))

    def copy_text(text):
        if text:
            QApplication.clipboard().setText(text)

    def copy_selected():
        rows = selected_row_numbers()
        copy_text(model.selected_rows_as_tsv(rows, include_headers=True, clean_copy=False))

    def copy_all():
        copy_text(
            model.all_as_tsv(
                include_headers=True,
                clean_copy=clean_copy,
                excluded_keys=copy_all_excluded_keys,
            )
        )

    def copy_column(column_index):
        rows = selected_row_numbers()
        if not rows:
            return
        key = model.columns[column_index].key
        values = [_clean_numeric_text(str(model.records[row].get(key, ""))) for row in rows]
        copy_text("\n".join(values))

    def copy_cell(row_index, column_index):
        if row_index < 0 or column_index < 0:
            return
        key = model.columns[column_index].key
        copy_text(_clean_numeric_text(str(model.records[row_index].get(key, ""))))

    def show_menu(pos):
        menu = QMenu(table)
        menu.addAction("Copy Selected Row(s)", copy_selected)
        menu.addAction("Copy All Data", copy_all)

        if allow_cell_column:
            index = table.indexAt(pos)
            if index.isValid():
                menu.addSeparator()
                header = model.columns[index.column()].header
                menu.addAction(f"Copy '{header}' Column (Selected Rows)", lambda: copy_column(index.column()))
                menu.addAction(f"Copy Cell: {index.data()}", lambda: copy_cell(index.row(), index.column()))

        menu.exec(table.viewport().mapToGlobal(pos))

    table.customContextMenuRequested.connect(show_menu)


def _clean_numeric_text(value):
    if value.replace(".", "", 1).replace("-", "", 1).isdigit() and "." in value:
        return value.replace(".", "")
    return value
