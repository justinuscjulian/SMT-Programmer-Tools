import os
import re
import shutil
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from services.errors import ServiceError


EXCEL_EXTENSIONS = (".xls", ".xlsx", ".xlsm", ".xlsb")
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XLSX_FILE_FORMAT = 51
COM_BUSY_HRESULTS = {-2147418111, -2147417846, -2147417845}


@dataclass
class InsertPointConfig:
    plan_path: str
    main_folder: str
    start_row: int
    end_row: int
    output_path: str


@dataclass
class InsertPointResult:
    output_path: str
    success_count: int
    error_count: int


def suggest_output_name():
    return f"DATA INSERT POINT ({datetime.now().strftime('%d-%m-%Y')}).xlsx"


def generate_insert_point(config: InsertPointConfig):
    _validate_config(config)
    win32, pythoncom = _load_win32()

    output_path = _normalize_output_path(config.output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    success = []
    errors = []
    excel = None
    plan_wb = None
    out_wb = None

    pythoncom.CoInitialize()
    try:
        try:
            excel = _excel_call(lambda: win32.DispatchEx("Excel.Application"))
        except Exception as exc:
            raise ServiceError(
                "Microsoft Excel tidak bisa dibuka. Pastikan Excel terinstall dan tidak sedang menampilkan dialog.",
                title="Excel tidak tersedia",
            ) from exc

        _configure_excel_for_background(excel)

        with tempfile.TemporaryDirectory(prefix="smt_insert_point_") as temp_dir:
            plan_wb = _open_workbook(
                excel,
                config.plan_path,
                ReadOnly=True,
            )
            ws_plan = _excel_call(lambda: plan_wb.ActiveSheet)

            for row in range(config.start_row, config.end_row + 1):
                part_num = str(_cell_value(ws_plan, row, "G") or "").strip()
                if not part_num:
                    continue

                model_name = str(_cell_value(ws_plan, row, "D") or "").strip()
                pcb_num = str(_cell_value(ws_plan, row, "I") or "").strip()

                if "." in model_name:
                    model_name = model_name.split(".")[0]

                num_only = extract_num_only(part_num)

                if not pcb_num:
                    errors.append([part_num, pcb_num, "", "PCB Part Number kosong di kolom I"])
                    continue

                if not num_only:
                    errors.append([part_num, pcb_num, "", "Part Number tidak mengandung angka untuk pencarian file"])
                    continue

                target_file = find_target_file(config.main_folder, pcb_num, num_only)

                if not target_file:
                    errors.append([part_num, pcb_num, "", f"File tidak ditemukan. Dicari pakai angka: {num_only}"])
                    continue

                wb_prog = None
                temp_path = None

                try:
                    wb_prog, temp_path = open_workbook_robust(excel, target_file, temp_dir)

                    try:
                        ws_dx = _excel_call(lambda: wb_prog.Worksheets("DX"))
                    except Exception:
                        errors.append([part_num, pcb_num, target_file, "Sheet 'DX' tidak ada"])
                        continue

                    insert_val = _range_value(ws_dx, "T9")

                    if insert_val is None or str(insert_val).strip() == "":
                        errors.append([part_num, pcb_num, target_file, "Cell T9 kosong"])
                    else:
                        success.append([model_name, part_num, pcb_num, insert_val])

                except Exception as exc:
                    errors.append([part_num, pcb_num, target_file, f"Gagal membuka/membaca file Excel: {exc}"])

                finally:
                    if wb_prog:
                        _close_workbook(wb_prog)

                    if temp_path and os.path.exists(temp_path):
                        os.remove(temp_path)

            _close_workbook(plan_wb)
            plan_wb = None

            out_wb = _excel_call(lambda: excel.Workbooks.Add())
            _write_output_workbook(out_wb, success, errors)

            if output_path.exists():
                output_path.unlink()

            _excel_call(lambda: out_wb.SaveAs(str(output_path), FileFormat=XLSX_FILE_FORMAT))
            _close_workbook(out_wb)
            out_wb = None

    finally:
        if plan_wb:
            _close_workbook(plan_wb)
        if out_wb:
            _close_workbook(out_wb)
        if excel:
            _restore_and_quit_excel(excel)
        pythoncom.CoUninitialize()

    return InsertPointResult(str(output_path), len(success), len(errors))


def extract_num_only(text):
    match = re.search(r"\d+", str(text))
    return match.group(0) if match else ""


def is_excel_file(filename):
    return filename.lower().endswith(EXCEL_EXTENSIONS)


def get_date_code_from_filename(filepath):
    base = Path(filepath).stem
    last6 = base[-6:]
    if last6.isdigit():
        return int(last6)

    modified = datetime.fromtimestamp(os.path.getmtime(filepath))
    return int(modified.strftime("%y%m%d"))


def find_target_file(main_folder, pcb_num, num_only):
    best_file = None
    best_date = -1

    for subfolder in Path(main_folder).iterdir():
        if not subfolder.is_dir():
            continue

        if pcb_num.lower() not in subfolder.name.lower():
            continue

        for file in subfolder.iterdir():
            if not file.is_file():
                continue

            if not is_excel_file(file.name):
                continue

            if num_only.lower() not in file.name.lower():
                continue

            file_date = get_date_code_from_filename(file)
            if file_date >= best_date:
                best_date = file_date
                best_file = str(file)

        if best_file:
            break

    return best_file


def open_workbook_robust(excel, source_path, temp_dir):
    try:
        if len(source_path) <= 200:
            return _open_workbook(excel, source_path), None
    except Exception:
        pass

    ext = Path(source_path).suffix
    temp_path = os.path.join(
        temp_dir,
        f"TEMP_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}{ext}",
    )
    shutil.copy2(source_path, temp_path)

    try:
        return _open_workbook(excel, temp_path), temp_path
    except Exception as exc:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise exc


def _write_output_workbook(out_wb, success, errors):
    ws_ok = _excel_call(lambda: out_wb.Worksheets(1))
    _excel_call(lambda: setattr(ws_ok, "Name", "INSERT POINT DATA"))
    _set_range_value(ws_ok, "A1:E1", [["No", "Model", "Part Number", "PCB Part Number", "Insert Point"]])

    for index, data in enumerate(success, start=1):
        _set_cell_value(ws_ok, index + 1, 1, index)
        _set_cell_value(ws_ok, index + 1, 2, data[0])
        _set_cell_value(ws_ok, index + 1, 3, data[1])
        _set_cell_value(ws_ok, index + 1, 4, data[2])
        _set_cell_value(ws_ok, index + 1, 5, data[3])

    ws_err = _excel_call(lambda: out_wb.Worksheets.Add(After=ws_ok))
    _excel_call(lambda: setattr(ws_err, "Name", "ERROR LOG"))
    _set_range_value(ws_err, "A1:E1", [["No", "Part Number", "PCB Part Number", "File Path", "Error Desc"]])

    for index, data in enumerate(errors, start=1):
        _set_cell_value(ws_err, index + 1, 1, index)
        _set_cell_value(ws_err, index + 1, 2, data[0])
        _set_cell_value(ws_err, index + 1, 3, data[1])
        _set_cell_value(ws_err, index + 1, 4, data[2])
        _set_cell_value(ws_err, index + 1, 5, data[3])

    _excel_call(lambda: ws_ok.Columns("A:E").AutoFit())
    _excel_call(lambda: ws_err.Columns("A:E").AutoFit())
    _excel_call(lambda: setattr(ws_err.Columns("E:E"), "ColumnWidth", 70))


def _open_workbook(excel, path, **kwargs):
    options = {
        "UpdateLinks": 0,
        "ReadOnly": True,
        "AddToMru": False,
        "IgnoreReadOnlyRecommended": True,
        "Notify": False,
    }
    options.update(kwargs)
    return _excel_call(lambda: excel.Workbooks.Open(path, **options))


def _cell_value(worksheet, row, column):
    return _excel_call(lambda: worksheet.Cells(row, column).Value)


def _set_cell_value(worksheet, row, column, value):
    _excel_call(lambda: setattr(worksheet.Cells(row, column), "Value", value))


def _range_value(worksheet, address):
    return _excel_call(lambda: worksheet.Range(address).Value)


def _set_range_value(worksheet, address, value):
    _excel_call(lambda: setattr(worksheet.Range(address), "Value", value))


def _excel_call(fn, attempts=80, delay=0.25):
    for attempt in range(attempts):
        try:
            return fn()
        except Exception as exc:
            if not _is_excel_busy_error(exc) or attempt == attempts - 1:
                raise
            time.sleep(delay)

    return None


def _is_excel_busy_error(exc):
    hresult = getattr(exc, "hresult", None)
    if hresult is None and getattr(exc, "args", None):
        hresult = exc.args[0]
    return hresult in COM_BUSY_HRESULTS


def _close_workbook(workbook):
    try:
        _excel_call(lambda: workbook.Close(SaveChanges=False))
    except Exception:
        pass


def _configure_excel_for_background(excel):
    for attr, value in (
        ("Visible", False),
        ("DisplayAlerts", False),
        ("ScreenUpdating", False),
        ("EnableEvents", False),
        ("Calculation", XL_CALCULATION_MANUAL),
    ):
        _set_excel_option(excel, attr, value)


def _set_excel_option(excel, attr, value):
    try:
        setattr(excel, attr, value)
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
        _excel_call(lambda: excel.Quit())
    except Exception:
        pass


def _validate_config(config):
    if not config.plan_path:
        raise ServiceError("File Excel PLAN belum dipilih.", title="Input belum lengkap")
    if not Path(config.plan_path).is_file():
        raise ServiceError(f"File Excel PLAN tidak ditemukan:\n{config.plan_path}", title="File tidak ditemukan")
    if not is_excel_file(config.plan_path):
        raise ServiceError("File PLAN harus berupa Excel (.xlsx, .xlsm, .xls, .xlsb).", title="Format tidak valid")

    if not config.main_folder:
        raise ServiceError("Folder Induk PCB belum dipilih.", title="Input belum lengkap")
    if not Path(config.main_folder).is_dir():
        raise ServiceError(f"Folder Induk PCB tidak ditemukan:\n{config.main_folder}", title="Folder tidak ditemukan")

    if config.start_row < 1 or config.end_row < 1:
        raise ServiceError("Start Row dan End Row minimal 1.", title="Range row tidak valid")
    if config.end_row < config.start_row:
        raise ServiceError("End Row tidak boleh lebih kecil dari Start Row.", title="Range row tidak valid")

    if not config.output_path:
        raise ServiceError("Lokasi output belum dipilih.", title="Input belum lengkap")


def _normalize_output_path(path):
    output_path = Path(path)
    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")
    return output_path


def _load_win32():
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError as exc:
        raise ServiceError(
            "Fitur Insert Point membutuhkan pywin32 dan Microsoft Excel terinstall.",
            title="Dependency belum lengkap",
        ) from exc
    return win32, pythoncom
