import os
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from services.errors import ServiceError


PLAN_TYPE_FIRST = "first"
PLAN_TYPE_SECOND = "second"
PLAN_TYPE_THIRD = "third"

EXCEL_EXTENSIONS = (".xls", ".xlsx", ".xlsm", ".xlsb")
XL_CALCULATION_AUTOMATIC = -4105
XL_NONE = -4142
XL_SOLID = 1
GREY_FILL = 11184814
BLACK_FONT = 0


@dataclass
class PlanConfig:
    plan_type: str
    previous_plan_path: str
    new_plan_path: str
    output_path: str


@dataclass
class PlanResult:
    output_path: str
    sheet_name: str
    matched_count: int
    not_found_count: int


@dataclass
class PlanBlock:
    key: str
    start_row: int
    end_row: int


def suggest_output_name(plan_type, previous_plan_path="", new_plan_path=""):
    suffix_by_type = {
        PLAN_TYPE_FIRST: "1ST",
        PLAN_TYPE_SECOND: "2ND",
        PLAN_TYPE_THIRD: "3RD",
    }
    suffix = suffix_by_type.get(plan_type, "1ST")

    if plan_type == PLAN_TYPE_FIRST:
        stem = Path(new_plan_path).stem if new_plan_path else "OHM PLAN"
        plan_name = re.sub(r"^\d+\.\s*", "", stem).strip() or stem
        return f"(RE_{datetime.now().strftime('%y%m%d')}_{suffix}) {plan_name}.xlsx"

    previous_name = Path(previous_plan_path).name if previous_plan_path else ""
    match = re.match(r"^\(RE_(\d{6})_(1ST|2ND|3RD)\)\s+(.+)\.xlsx$", previous_name, re.IGNORECASE)
    if match:
        date_code, _, plan_name = match.groups()
        return f"(RE_{date_code}_{suffix}) {plan_name}.xlsx"

    stem = Path(new_plan_path).stem if new_plan_path else "OHM PLAN"
    plan_name = re.sub(r"^\(RE_\d{6}_(?:1ST|2ND|3RD)\)\s+", "", stem, flags=re.IGNORECASE)
    plan_name = re.sub(r"^\d+\.\s*", "", plan_name).strip() or stem
    return f"(RE_{datetime.now().strftime('%y%m%d')}_{suffix}) {plan_name}.xlsx"


def generate_plan(config: PlanConfig):
    _validate_config(config)
    win32, pythoncom = _load_win32()

    output_path = _normalize_output_path(config.output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        if output_path.exists():
            output_path.unlink()
        shutil.copy2(config.new_plan_path, output_path)
    except OSError as exc:
        raise ServiceError(
            f"Output tidak bisa ditulis. Pastikan file tujuan tidak sedang terbuka:\n{output_path}",
            title="Output terkunci",
        ) from exc

    excel = None
    previous_wb = None
    output_wb = None

    pythoncom.CoInitialize()
    try:
        try:
            excel = win32.DispatchEx("Excel.Application")
        except Exception as exc:
            raise ServiceError(
                "Microsoft Excel tidak bisa dibuka. Pastikan Excel terinstall dan tidak sedang menampilkan dialog.",
                title="Excel tidak tersedia",
            ) from exc

        _configure_excel_for_background(excel)

        previous_wb = excel.Workbooks.Open(
            str(Path(config.previous_plan_path).resolve()),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
        )
        output_wb = excel.Workbooks.Open(
            str(output_path.resolve()),
            UpdateLinks=0,
            ReadOnly=False,
            AddToMru=False,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
        )

        previous_ws = previous_wb.ActiveSheet
        output_ws = output_wb.ActiveSheet

        if config.plan_type == PLAN_TYPE_FIRST:
            _apply_first_plan_adjustments(output_ws)
            matched_count, not_found_count = _process_first_plan(output_ws, previous_ws)
        else:
            _apply_next_plan_adjustments(output_ws)
            matched_count, not_found_count = _process_next_plan(output_ws, previous_ws)

        output_wb.Save()
        sheet_name = output_ws.Name

    finally:
        if output_wb:
            _close_workbook(output_wb)
        if previous_wb:
            _close_workbook(previous_wb)
        if excel:
            _restore_and_quit_excel(excel)
        pythoncom.CoUninitialize()

    return PlanResult(str(output_path), sheet_name, matched_count, not_found_count)


def _process_first_plan(output_ws, previous_ws):
    output_blocks = _find_plan_blocks(output_ws)
    previous_lookup = _build_lookup_by_block(previous_ws, key_col=3, value_col=22)
    matched_count = 0
    not_found_count = 0

    for block in output_blocks:
        lookup = previous_lookup.get(block.key, {})
        for row in range(block.start_row, block.end_row + 1):
            key = _cell_text(output_ws.Cells(row, 3).Value)
            if not key:
                continue

            matched = lookup.get(key.upper())
            target_cell = output_ws.Cells(row, 22)
            if matched is None:
                target_cell.Value = "#N/A"
                not_found_count += 1
                continue

            target_cell.Value = matched["value"]
            _set_grey_fill(target_cell)
            _set_grey_fill(output_ws.Cells(row, 7))
            matched_count += 1

    _clear_column_borders(output_ws, 22)
    return matched_count, not_found_count


def _process_next_plan(output_ws, previous_ws):
    output_blocks = _find_plan_blocks(output_ws)
    previous_lookup = _build_lookup_by_block(previous_ws, key_col=3, value_col=22, extra_cols=(7, 16))
    matched_count = 0
    not_found_count = 0

    for block in output_blocks:
        lookup = previous_lookup.get(block.key, {})
        for row in range(block.start_row, block.end_row + 1):
            key = _cell_text(output_ws.Cells(row, 3).Value)
            if not key:
                continue

            matched = lookup.get(key.upper())
            target_cell = output_ws.Cells(row, 22)
            if matched is None:
                target_cell.Value = "#N/A"
                not_found_count += 1
                continue

            target_cell.Value = matched["value"]
            _copy_fill(matched["value_cell"], target_cell)
            _copy_fill(matched["extra_cells"][7], output_ws.Cells(row, 7))

            s2_source = matched["extra_cells"][16]
            s2_target = output_ws.Cells(row, 16)
            s2_target.Value = s2_source.Value
            _copy_fill(s2_source, s2_target)
            matched_count += 1

    _set_column_font_color(output_ws, 16, BLACK_FONT)
    _clear_column_borders(output_ws, 22)
    return matched_count, not_found_count


def _apply_first_plan_adjustments(ws):
    _clean_plan_sheet(ws)
    _insert_first_plan_program_column(ws)
    _clear_target_column(ws, 22)
    _apply_first_plan_column_layout(ws)
    _select_cell(ws, "V4")


def _apply_next_plan_adjustments(ws):
    _clean_plan_sheet(ws)
    _insert_first_plan_program_column(ws)
    _clear_target_column(ws, 22)
    _apply_first_plan_column_layout(ws)
    _select_cell(ws, "V4")


def _insert_first_plan_program_column(ws):
    ws.Columns(24).Insert(Shift=-4161, CopyOrigin=0)
    ws.Columns(24).Cut()
    ws.Columns(22).Insert(Shift=-4161)


def _clean_plan_sheet(ws):
    try:
        ws.Calculate()
    except Exception:
        pass

    used_range = ws.UsedRange
    used_range.Value = used_range.Value

    font = ws.Cells.Font
    font.Name = "Calibri"
    font.Size = 10
    font.Strikethrough = False
    font.Superscript = False
    font.Subscript = False
    font.Underline = -4142
    font.TintAndShade = 0
    font.ThemeFont = 2


def _apply_first_plan_column_layout(ws):
    try:
        ws.Columns(5).EntireColumn.Hidden = True
        ws.Columns(21).EntireColumn.Hidden = True
        ws.Columns(23).EntireColumn.Hidden = True
        ws.Columns(24).EntireColumn.Hidden = True
        ws.Columns(20).ColumnWidth = 1.5
        ws.Columns(22).ColumnWidth = 105
        ws.Columns(16).ColumnWidth = 17
        ws.Columns(7).ColumnWidth = 11
        ws.Columns(22).EntireColumn.Hidden = False
        _clear_column_fill_and_borders(ws, 22)
    except Exception:
        pass

    try:
        ws.Parent.Windows(1).Zoom = 110
    except Exception:
        pass


def _clear_target_column(ws, column):
    used_range = ws.UsedRange
    first_row = used_range.Row
    last_row = used_range.Row + used_range.Rows.Count - 1
    target_range = ws.Range(ws.Cells(first_row, column), ws.Cells(last_row, column))
    target_range.ClearContents()
    target_range.Interior.Pattern = XL_NONE
    target_range.Borders.LineStyle = XL_NONE


def _clear_column_fill_and_borders(ws, column):
    target = ws.Columns(column)
    target.Interior.Pattern = XL_NONE
    _clear_column_borders(ws, column)


def _clear_column_borders(ws, column):
    target = ws.Columns(column)
    target.Borders.LineStyle = XL_NONE
    for border_index in range(7, 13):
        target.Borders(border_index).LineStyle = XL_NONE

    if column > 1:
        ws.Columns(column - 1).Borders(10).LineStyle = XL_NONE
    ws.Columns(column + 1).Borders(7).LineStyle = XL_NONE

    used_range = ws.UsedRange
    first_row = used_range.Row
    last_row = used_range.Row + used_range.Rows.Count - 1
    used_column_range = ws.Range(ws.Cells(first_row, column), ws.Cells(last_row, column))
    used_column_range.Borders.LineStyle = XL_NONE
    for border_index in range(7, 13):
        used_column_range.Borders(border_index).LineStyle = XL_NONE


def _set_column_font_color(ws, column, color):
    ws.Columns(column).Font.Color = color


def _set_grey_fill(cell):
    cell.Interior.Pattern = XL_SOLID
    cell.Interior.Color = GREY_FILL


def _select_cell(ws, address):
    try:
        ws.Range(address).Select()
    except Exception:
        pass


def _find_plan_blocks(ws):
    used_range = ws.UsedRange
    first_row = used_range.Row
    last_row = used_range.Row + used_range.Rows.Count - 1
    blocks = []
    row = first_row

    while row <= last_row:
        if _cell_text(ws.Cells(row, 1).Value).upper() == "LINE":
            start_row = row + 1
            if not _cell_text(ws.Cells(start_row, 1).Value):
                row += 1
                continue

            end_row = start_row
            while end_row + 1 <= last_row and _cell_text(ws.Cells(end_row + 1, 1).Value):
                end_row += 1

            block_key = _line_key(ws.Cells(start_row, 1).Value)
            if block_key:
                blocks.append(PlanBlock(block_key, start_row, end_row))
            row = end_row + 1
        row += 1

    if not blocks:
        raise ServiceError("Blok PLAN tidak ditemukan pada active sheet file PLAN baru.", title="Format PLAN tidak valid")

    return blocks


def _build_lookup_by_block(ws, key_col, value_col, extra_cols=()):
    lookup_by_block = {}
    for block in _find_plan_blocks(ws):
        lookup = {}
        for row in range(block.start_row, block.end_row + 1):
            key = _cell_text(ws.Cells(row, key_col).Value)
            if not key:
                continue

            normalized = key.upper()
            if normalized in lookup:
                continue

            value_cell = ws.Cells(row, value_col)
            lookup[normalized] = {
                "value": value_cell.Value,
                "value_cell": value_cell,
                "extra_cells": {col: ws.Cells(row, col) for col in extra_cols},
            }
        lookup_by_block[block.key] = lookup
    return lookup_by_block


def _line_key(value):
    text = _cell_text(value).upper()
    match = re.match(r"^(\d+\s*LINE|LGERC)", text)
    if not match:
        return ""
    return re.sub(r"\s+", "", match.group(1))


def _copy_fill(source_cell, target_cell):
    try:
        if source_cell.Interior.Pattern == XL_NONE:
            target_cell.Interior.Pattern = XL_NONE
        else:
            target_cell.Interior.Pattern = XL_SOLID
            target_cell.Interior.Color = source_cell.Interior.Color
    except Exception:
        pass


def _cell_text(value):
    if value is None:
        return ""
    return str(value).strip()


def _configure_excel_for_background(excel):
    for attr, value in (
        ("Visible", False),
        ("DisplayAlerts", False),
        ("ScreenUpdating", False),
        ("EnableEvents", False),
        ("Calculation", XL_CALCULATION_AUTOMATIC),
    ):
        _set_excel_option(excel, attr, value)


def _set_excel_option(excel, attr, value):
    try:
        setattr(excel, attr, value)
    except Exception:
        pass


def _close_workbook(workbook):
    try:
        workbook.Close(SaveChanges=False)
    except Exception:
        pass


def _restore_and_quit_excel(excel):
    for attr, value in (
        ("Calculation", XL_CALCULATION_AUTOMATIC),
        ("ScreenUpdating", True),
        ("EnableEvents", True),
        ("DisplayAlerts", True),
    ):
        _set_excel_option(excel, attr, value)

    try:
        excel.Quit()
    except Exception:
        pass


def _validate_config(config):
    if config.plan_type not in {PLAN_TYPE_FIRST, PLAN_TYPE_SECOND, PLAN_TYPE_THIRD}:
        raise ServiceError("Tipe PLAN tidak valid.", title="Input tidak valid")

    for path, label in (
        (config.previous_plan_path, "PLAN sebelumnya"),
        (config.new_plan_path, "PLAN baru"),
    ):
        if not path:
            raise ServiceError(f"{label} belum dipilih.", title="Input belum lengkap")
        if not Path(path).is_file():
            raise ServiceError(f"{label} tidak ditemukan:\n{path}", title="File tidak ditemukan")
        if not _is_excel_file(path):
            raise ServiceError(f"{label} harus berupa Excel (.xlsx, .xlsm, .xls, .xlsb).", title="Format tidak valid")

    if not config.output_path:
        raise ServiceError("Lokasi output belum dipilih.", title="Input belum lengkap")

    output_path = _normalize_output_path(config.output_path)
    input_paths = {Path(config.previous_plan_path).resolve(), Path(config.new_plan_path).resolve()}
    if output_path.resolve() in input_paths:
        raise ServiceError("Lokasi output tidak boleh sama dengan file input.", title="Output tidak valid")


def _normalize_output_path(path):
    output_path = Path(path)
    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")
    return output_path


def _is_excel_file(path):
    return str(path).lower().endswith(EXCEL_EXTENSIONS)


def _load_win32():
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError as exc:
        raise ServiceError(
            "Fitur PLAN membutuhkan pywin32 dan Microsoft Excel terinstall.",
            title="Dependency belum lengkap",
        ) from exc
    return win32, pythoncom
