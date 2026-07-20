import pandas as pd
from dataclasses import dataclass
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook

from services.errors import ServiceError
from services.model_feeder_group_service import _scan_models, _emit_progress
from services.common_feeder_reuse_service import _normalize_output_path, _style_sheet


@dataclass
class Table7FeederConfig:
    source_folder: str
    target_pcb_list: list = None
    table7_ref_file: str = ""


@dataclass
class Table7FeederResult:
    pcb_rows: list[dict]
    total_files: int
    read_files: int
    skipped_files: list[str]
    model_count: int


def get_table7_components(ref_file: str) -> set:
    try:
        df = pd.read_excel(ref_file)
        if df.empty:
            raise ServiceError("File komponen Table 7 kosong.", title="Data Kosong")
        # Assume first column has the data
        components = set(df.iloc[:, 0].dropna().astype(str).str.strip().str.upper())
        if not components:
            raise ServiceError("Tidak ada komponen valid di file.", title="Data Kosong")
        return components
    except Exception as exc:
        raise ServiceError(f"Gagal membaca file komponen Table 7: {exc}", title="Error Baca File")


def analyze_table7_feeders(config: Table7FeederConfig, progress_callback=None) -> Table7FeederResult:
    if not config.source_folder or not Path(config.source_folder).is_dir():
        raise ServiceError("Folder Induk PCB tidak valid.", title="Input Error")
    if not config.table7_ref_file or not Path(config.table7_ref_file).is_file():
        raise ServiceError("File referensi Table 7 tidak valid.", title="Input Error")

    _emit_progress(progress_callback, 0, "Membaca komponen Table 7...")
    table7_components = get_table7_components(config.table7_ref_file)

    _emit_progress(progress_callback, 5, "Scanning PCB folders...")
    models, total_files, read_files, skipped_files = _scan_models(
        config.source_folder, config.target_pcb_list, progress_callback
    )

    if not models:
        return Table7FeederResult([], total_files, read_files, skipped_files, 0)

    _emit_progress(progress_callback, 96, "Menganalisa komponen Table 7 per PCB...")
    pcb_rows = []
    
    for key, model in models.items():
        # Get intersection of model's components and Table 7 components
        used_table7_parts = []
        for comp_key, comp_val in model.components.items():
            if comp_key.upper() in table7_components or comp_val.upper() in table7_components:
                used_table7_parts.append(comp_val.upper())
                
        # Remove duplicates while preserving order if possible (they are already unique mostly)
        used_table7_parts = list(dict.fromkeys(used_table7_parts))
        used_table7_parts.sort()
        
        comp_count = len(used_table7_parts)
        if comp_count == 0:
            status = "NO TABLE 7 PARTS"
        elif comp_count <= 30:
            status = "OK"
        else:
            status = "OVERLOAD (> 30)"
            
        # Format slot assignment string
        slot_assignments = []
        for idx, part in enumerate(used_table7_parts, start=1):
            slot_assignments.append(f"Slot {idx}: {part}")
        
        pcb_rows.append({
            "pcb_part_number": model.pcb_part_number,
            "status": status,
            "table7_part_count": comp_count,
            "excel_file_count": len(model.source_files),
            "members": "; ".join(model.source_files),
            "slot_assignments": "\n".join(slot_assignments) if slot_assignments else "-",
            "_raw_parts": used_table7_parts,
        })
        
    _emit_progress(progress_callback, 100, "Selesai")
    
    return Table7FeederResult(
        pcb_rows=pcb_rows,
        total_files=total_files,
        read_files=read_files,
        skipped_files=skipped_files,
        model_count=len(models),
    )


def suggest_table7_export_name():
    return f"Table7_Fix_Feeder_{datetime.now().strftime('%y%m%d')}.xlsx"


def export_table7_result(result: Table7FeederResult, output_path: str):
    if result is None or not result.pcb_rows:
        raise ServiceError("Belum ada hasil analisa untuk diexport.", title="Data kosong")

    output = _normalize_output_path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    
    # 1. Summary Sheet
    summary_sheet = workbook.active
    summary_sheet.title = "Table 7 Fix Feeder Summary"
    columns = [
        ("pcb_part_number", "PCB Part Number"),
        ("status", "Status"),
        ("table7_part_count", "Total Parts T7"),
        ("excel_file_count", "Total Program Excel"),
        ("members", "Program Excel"),
    ]
    summary_sheet.append([header for _, header in columns])
    for row in result.pcb_rows:
        summary_sheet.append([row.get(key, "") for key, _ in columns])
    _style_sheet(summary_sheet)
    
    # 2. Detailed Slots Sheet
    slot_sheet = workbook.create_sheet("Table 7 Slots Detail")
    slot_sheet.append(["PCB Part Number", "Status", "Total Parts T7"] + [f"Slot {i}" for i in range(1, 31)])
    for row in result.pcb_rows:
        out_row = [row["pcb_part_number"], row["status"], row["table7_part_count"]]
        parts = row["_raw_parts"]
        # Fill slots 1 to 30
        for i in range(30):
            if i < len(parts):
                out_row.append(parts[i])
            else:
                out_row.append("")
        slot_sheet.append(out_row)
    _style_sheet(slot_sheet)
    
    # 3. Log Sheet
    log_sheet = workbook.create_sheet("Scan Log")
    log_rows = [
        ("Excel files found", result.total_files),
        ("Files read", result.read_files),
        ("PCB folders analyzed", result.model_count),
        ("Skipped/error files", len(result.skipped_files)),
    ]
    log_sheet.append(["Item", "Value"])
    for item, value in log_rows:
        log_sheet.append([item, value])
        
    if result.skipped_files:
        log_sheet.append([])
        log_sheet.append(["Skipped/error detail", ""])
        for skipped in result.skipped_files:
            log_sheet.append([skipped, ""])
    _style_sheet(log_sheet)

    workbook.save(output)
    return str(output)
