from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QAbstractItemView, QHeaderView, QMenu, QTableView


def configure_table(table: QTableView, model):
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
    table.setShowGrid(False)

    for idx, column in enumerate(getattr(model, "columns", [])):
        table.setColumnWidth(idx, column.width)


def install_copy_menu(table, model, clean_copy=False, allow_cell_column=False):
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
        copy_text(model.all_as_tsv(include_headers=True, clean_copy=clean_copy))

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
