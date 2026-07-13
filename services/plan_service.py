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
    history_folder_path: str = ""


@dataclass
class PlanResult:
    output_path: str
    sheet_name: str
    matched_count: int
    not_found_count: int
    history_found_count: int = 0
    history_other_line_count: int = 0
    history_not_found_count: int = 0


@dataclass
class PlanBlock:
    key: str
    start_row: int
    end_row: int


@dataclass
class PlanProcessResult:
    matched_count: int = 0
    not_found_count: int = 0
    history_found_count: int = 0
    history_other_line_count: int = 0
    history_not_found_count: int = 0


@dataclass
class HistoryProgramMatch:
    value: str
    status: str


HISTORY_STATUS_FOUND = "found"
HISTORY_STATUS_OTHER_LINE = "other_line"
HISTORY_STATUS_NOT_FOUND = "not_found"
HISTORY_OTHER_LINE_TEXT = "Program ada di LINE lain"
HISTORY_NOT_FOUND_TEXT = "History program tidak ditemukan"


@dataclass
class PlanLayout:
    col_line: int = 1
    col_wo_supply: int = 3
    col_smt_pn: int = 5
    col_dms_pn: int = 7
    col_pcb: int = 9
    col_s2: int = 16
    col_dms_fk: int = 21
    col_target: int = 22
    col_device: int = 23
    col_extra_cut: int = 24
    header_row: int = 3

    @classmethod
    def analyze(cls, ws, header_row=3):
        layout = cls(header_row=header_row)
        headers = {}
        for col in range(1, 100):
            val = _cell_text(ws.Cells(header_row, col).Value).upper()
            if val:
                headers[val] = col
        
        def resolve(aliases, default):
            for alias in aliases:
                for h, c in headers.items():
                    if alias in h:
                        return c
            return default

        layout.col_line = resolve(["LINE"], 1)
        layout.col_wo_supply = resolve(["WO SUPPLY", "WO", "WORK ORDER"], 3)
        layout.col_smt_pn = resolve(["SMT P/N", "SMT PN", "SMT PART"], 5)
        layout.col_dms_pn = resolve(["DMS P/N", "DMS PN", "DMS PART"], 7)
        layout.col_pcb = resolve(["PCB"], 9)
        layout.col_s2 = resolve(["S-2", "S2", "REMARK"], 16)
        layout.col_dms_fk = resolve(["DMS F/K", "DMS FK"], 21)
        layout.col_device = resolve(["DEVICE"], 23)
        layout.col_target = layout.col_device if layout.col_device > 0 else 22
        
        # We assume the column to cut is the one right after DEVICE
        layout.col_extra_cut = layout.col_device + 1 if layout.col_device > 0 else 24
        return layout

    def insert_target_column(self, ws):
        # Emulate the exact original behavior but using mapped columns
        ws.Columns(self.col_extra_cut).Insert(Shift=-4161, CopyOrigin=0)
        ws.Columns(self.col_extra_cut).Cut()
        ws.Columns(self.col_target).Insert(Shift=-4161)
        
        insert_at = self.col_target
        # Adjust indices for columns that shifted
        if self.col_line >= insert_at: self.col_line += 1
        if self.col_wo_supply >= insert_at: self.col_wo_supply += 1
        if self.col_smt_pn >= insert_at: self.col_smt_pn += 1
        if self.col_dms_pn >= insert_at: self.col_dms_pn += 1
        if self.col_pcb >= insert_at: self.col_pcb += 1
        if self.col_s2 >= insert_at: self.col_s2 += 1
        if self.col_dms_fk >= insert_at: self.col_dms_fk += 1
        if self.col_device >= insert_at: self.col_device += 1
        if self.col_extra_cut >= insert_at: self.col_extra_cut += 1


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


class HistoryProgramIndex:
    def __init__(self, folder_path):
        self.files = []
        with os.scandir(folder_path) as entries:
            for entry in entries:
                if entry.is_file():
                    self.files.append((entry.name, entry.name.upper()))
        self.files.sort(key=lambda item: item[0].upper())

    def find(self, raw_part_number, pcb_number, line_id):
        part_digits = _digits_only(raw_part_number)
        pcb_text = _cell_text(pcb_number)

        if not part_digits or not pcb_text:
            return HistoryProgramMatch(HISTORY_NOT_FOUND_TEXT, HISTORY_STATUS_NOT_FOUND)

        part_token = part_digits.upper()
        pcb_token = pcb_text.upper()
        line_marker = f"(INI{line_id})".upper() if line_id else ""

        best_match = ""
        old_match = ""
        found_in_other_line = False
        any_file_found = False

        for file_name, upper_name in self.files:
            if part_token not in upper_name or pcb_token not in upper_name:
                continue

            any_file_found = True
            is_old = "(OLD)" in upper_name

            if line_marker:
                if line_marker in upper_name:
                    if is_old:
                        if not old_match:
                            old_match = file_name
                    else:
                        best_match = file_name
                        break
                else:
                    found_in_other_line = True
            elif is_old:
                if not old_match:
                    old_match = file_name
            elif not best_match:
                best_match = file_name

        final_file = best_match or old_match
        if final_file:
            return HistoryProgramMatch(final_file, HISTORY_STATUS_FOUND)
        if found_in_other_line:
            return HistoryProgramMatch(HISTORY_OTHER_LINE_TEXT, HISTORY_STATUS_OTHER_LINE)
        if not any_file_found:
            return HistoryProgramMatch(HISTORY_NOT_FOUND_TEXT, HISTORY_STATUS_NOT_FOUND)

        return HistoryProgramMatch(HISTORY_NOT_FOUND_TEXT, HISTORY_STATUS_NOT_FOUND)


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
        history_index = HistoryProgramIndex(config.history_folder_path)

        prev_layout = PlanLayout.analyze(previous_ws)
        out_layout = PlanLayout.analyze(output_ws)

        if config.plan_type == PLAN_TYPE_FIRST:
            _apply_first_plan_adjustments(output_ws, out_layout)
            process_result = _process_first_plan(output_ws, previous_ws, history_index, out_layout, prev_layout)
        else:
            _apply_next_plan_adjustments(output_ws, out_layout)
            process_result = _process_next_plan(output_ws, previous_ws, history_index, out_layout, prev_layout)

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

    return PlanResult(
        str(output_path),
        sheet_name,
        process_result.matched_count,
        process_result.not_found_count,
        process_result.history_found_count,
        process_result.history_other_line_count,
        process_result.history_not_found_count,
    )


def _process_first_plan(output_ws, previous_ws, history_index, out_layout, prev_layout):
    output_blocks = _find_plan_blocks(output_ws, out_layout)
    previous_lookup = _build_lookup_by_block(
        previous_ws, 
        prev_layout, 
        key_col=prev_layout.col_wo_supply, 
        value_col=prev_layout.col_target, 
        extra_cols=(prev_layout.col_dms_pn, prev_layout.col_s2)
    )
    result = PlanProcessResult()

    for block in output_blocks:
        lookup = previous_lookup.get(block.key, {})
        for row in range(block.start_row, block.end_row + 1):
            key = _cell_text(output_ws.Cells(row, out_layout.col_wo_supply).Value)
            if not key:
                continue

            matched = lookup.get(key.upper())
            target_cell = output_ws.Cells(row, out_layout.col_target)
            if matched is None:
                result.not_found_count += 1
                _apply_history_fallback(output_ws, row, target_cell, history_index, result, out_layout)
                continue

            target_cell.Value = matched["value"]
            
            prev_col_g = matched["extra_cells"][prev_layout.col_dms_pn]
            try:
                is_transparent = (prev_col_g.Interior.Pattern == XL_NONE)
            except Exception:
                is_transparent = False

            if is_transparent:
                _copy_fill(matched["value_cell"], target_cell)
                _copy_fill(prev_col_g, output_ws.Cells(row, out_layout.col_dms_pn))
                
                prev_col_p = matched["extra_cells"][prev_layout.col_s2]
                target_col_p = output_ws.Cells(row, out_layout.col_s2)
                target_col_p.Value = prev_col_p.Value
                _copy_fill(prev_col_p, target_col_p)
            else:
                _set_grey_fill(target_cell)
                _set_grey_fill(output_ws.Cells(row, out_layout.col_dms_pn))
                
            result.matched_count += 1

    _clear_column_borders(output_ws, out_layout.col_target)
    return result


def _process_next_plan(output_ws, previous_ws, history_index, out_layout, prev_layout):
    output_blocks = _find_plan_blocks(output_ws, out_layout)
    previous_lookup = _build_lookup_by_block(
        previous_ws, 
        prev_layout, 
        key_col=prev_layout.col_wo_supply, 
        value_col=prev_layout.col_target, 
        extra_cols=(prev_layout.col_dms_pn, prev_layout.col_s2)
    )
    result = PlanProcessResult()

    for block in output_blocks:
        lookup = previous_lookup.get(block.key, {})
        for row in range(block.start_row, block.end_row + 1):
            key = _cell_text(output_ws.Cells(row, out_layout.col_wo_supply).Value)
            if not key:
                continue

            matched = lookup.get(key.upper())
            target_cell = output_ws.Cells(row, out_layout.col_target)
            if matched is None:
                result.not_found_count += 1
                _apply_history_fallback(output_ws, row, target_cell, history_index, result, out_layout)
                continue

            target_cell.Value = matched["value"]
            _copy_fill(matched["value_cell"], target_cell)
            _copy_fill(matched["extra_cells"][prev_layout.col_dms_pn], output_ws.Cells(row, out_layout.col_dms_pn))

            s2_source = matched["extra_cells"][prev_layout.col_s2]
            s2_target = output_ws.Cells(row, out_layout.col_s2)
            s2_target.Value = s2_source.Value
            _copy_fill(s2_source, s2_target)
            result.matched_count += 1

    _set_column_font_color(output_ws, out_layout.col_s2, BLACK_FONT)
    _clear_column_borders(output_ws, out_layout.col_target)
    return result


def _apply_history_fallback(output_ws, row, target_cell, history_index, result, layout):
    history_match = history_index.find(
        output_ws.Cells(row, layout.col_dms_pn).Value,
        output_ws.Cells(row, layout.col_pcb).Value,
        _history_line_id(output_ws.Cells(row, layout.col_line).Value),
    )

    target_cell.Value = history_match.value
    _clear_cell_fill_and_borders(target_cell)

    if history_match.status == HISTORY_STATUS_FOUND:
        result.history_found_count += 1
    elif history_match.status == HISTORY_STATUS_OTHER_LINE:
        result.history_other_line_count += 1
    else:
        result.history_not_found_count += 1


def _apply_first_plan_adjustments(ws, layout):
    _clean_plan_sheet(ws)
    layout.insert_target_column(ws)
    _clear_target_column(ws, layout.col_target)
    _apply_first_plan_column_layout(ws, layout)
    # Convert V4 to proper column char if possible, or just ignore for now since it's just selection
    _select_cell(ws, "A1")


def _apply_next_plan_adjustments(ws, layout):
    _clean_plan_sheet(ws)
    layout.insert_target_column(ws)
    _clear_target_column(ws, layout.col_target)
    _apply_first_plan_column_layout(ws, layout)
    _select_cell(ws, "A1")


# This is now handled by PlanLayout.insert_target_column


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


def _apply_first_plan_column_layout(ws, layout):
    try:
        ws.Columns(layout.col_smt_pn).EntireColumn.Hidden = True
        ws.Columns(layout.col_dms_fk).EntireColumn.Hidden = True
        ws.Columns(layout.col_device).EntireColumn.Hidden = True
        ws.Columns(layout.col_extra_cut).EntireColumn.Hidden = True
        
        # Col 20 was hardcoded, let's just make it col_target - 2 if we don't know it,
        # but realistically the user just wants the widths of the target and its neighbors fixed.
        # If we can't find col 20 dynamically, we just hide it by index. But 20 is typically before DMS F/K.
        ws.Columns(layout.col_target - 2).ColumnWidth = 1.5 if layout.col_target >= 3 else 1.5
        
        ws.Columns(layout.col_target).ColumnWidth = 105
        ws.Columns(layout.col_s2).ColumnWidth = 17
        ws.Columns(layout.col_dms_pn).ColumnWidth = 11
        ws.Columns(layout.col_target).EntireColumn.Hidden = False
        _clear_column_fill_and_borders(ws, layout.col_target)
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


def _clear_cell_fill_and_borders(cell):
    cell.Interior.Pattern = XL_NONE
    cell.Borders.LineStyle = XL_NONE
    for border_index in range(7, 13):
        cell.Borders(border_index).LineStyle = XL_NONE


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


def _find_plan_blocks(ws, layout):
    used_range = ws.UsedRange
    first_row = used_range.Row
    last_row = used_range.Row + used_range.Rows.Count - 1
    blocks = []
    row = first_row

    while row <= last_row:
        if _cell_text(ws.Cells(row, layout.col_line).Value).upper() == "LINE":
            start_row = row + 1
            if not _cell_text(ws.Cells(start_row, layout.col_line).Value):
                row += 1
                continue

            end_row = start_row
            while end_row + 1 <= last_row and _cell_text(ws.Cells(end_row + 1, layout.col_line).Value):
                end_row += 1

            block_key = _line_key(ws.Cells(start_row, layout.col_line).Value)
            if block_key:
                blocks.append(PlanBlock(block_key, start_row, end_row))
            row = end_row + 1
        row += 1

    if not blocks:
        raise ServiceError("Blok PLAN tidak ditemukan pada active sheet file PLAN baru.", title="Format PLAN tidak valid")

    return blocks


def _build_lookup_by_block(ws, layout, key_col, value_col, extra_cols=()):
    lookup_by_block = {}
    for block in _find_plan_blocks(ws, layout):
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


def _history_line_id(value):
    text = _cell_text(value)
    if not text:
        return ""

    line_match = re.search(r"^(.*?)\s*LINE\b", text, flags=re.IGNORECASE)
    if line_match:
        prefix = re.sub(r"\s+", "", line_match.group(1).strip())
        if prefix:
            return prefix.upper()

    block_key = _line_key(text)
    if block_key.endswith("LINE"):
        return block_key[:-4]
    return block_key


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


def _digits_only(value):
    return "".join(char for char in _cell_text(value) if char.isdigit())


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

    if not config.history_folder_path:
        raise ServiceError("Folder history belum dipilih.", title="Input belum lengkap")
    if not Path(config.history_folder_path).is_dir():
        raise ServiceError(
            f"Folder history tidak ditemukan:\n{config.history_folder_path}",
            title="Folder tidak ditemukan",
        )

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
