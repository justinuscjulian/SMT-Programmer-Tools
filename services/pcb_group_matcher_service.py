import csv
import math
import re
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from services.common_feeder_reuse_service import _clean_part_values, _part_key, _part_text, _read_bom_parts
from services.errors import ServiceError
from services.io_helpers import scan_recursive_files

MODE_BOM_FILE = "bom_file"
MODE_PROGRAM_FOLDER = "program_folder"
MODE_PROGRAM_EXCEL = "program_excel"

EXCEL_EXTENSIONS = (".xls", ".xlsx", ".xlsm", ".xlsb")
TEXT_EXTENSIONS = (".csv", ".tsv", ".txt")


@dataclass
class PcbGroupMatcherConfig:
    import_mode: str
    pcb_source_path: str
    fix_feeder_group_path: str
    output_path: str = ""


@dataclass
class GroupMatchItem:
    rank: int
    group_name: str
    member_pcbs: list[str]
    total_pcb_parts: int
    matched_count: int
    missing_count: int
    match_rate_percent: float
    status_recommendation: str
    recommendation_note: str
    matched_details: list[dict] = field(default_factory=list)
    missing_details: list[str] = field(default_factory=list)


@dataclass
class PcbGroupMatcherResult:
    output_path: str
    pcb_name: str
    import_mode: str
    pcb_parts_count: int
    unique_parts: list[str]
    group_matches: list[GroupMatchItem]
    best_match: GroupMatchItem


def suggest_output_name(pcb_name="NEW_PCB"):
    clean_name = re.sub(r'[\\/*?:"<>|]', '_', pcb_name).strip() or "NEW_PCB"
    return f"PCB_Group_Matcher_{clean_name}_{datetime.now().strftime('%y%m%d')}.xlsx"


def analyze_pcb_group_matcher(config: PcbGroupMatcherConfig, progress_callback=None):
    _validate_config(config)

    _emit_progress(progress_callback, 5, "Membaca komponen PCB baru...")
    pcb_parts, pcb_name = _extract_new_pcb_parts(config.import_mode, config.pcb_source_path)
    if not pcb_parts:
        raise ServiceError(
            "Tidak ditemukan Part Number komponen yang valid dari input PCB baru yang diberikan.",
            title="Komponen kosong",
        )

    _emit_progress(progress_callback, 25, "Membaca data Fix Feeder Groups dari Excel...")
    group_parts_map, group_pcbs_map = _load_fix_feeder_groups(config.fix_feeder_group_path)
    if not group_parts_map:
        raise ServiceError(
            "File Fix Feeder Group Excel tidak berisi data sheet group yang valid.",
            title="Data Fix Feeder kosong",
        )

    _emit_progress(progress_callback, 55, "Melakukan komparasi kecocokan PCB ke seluruh Group...")
    raw_matches = []
    total_groups = len(group_parts_map)
    total_parts = len(pcb_parts)

    for idx, (group_name, group_parts) in enumerate(group_parts_map.items(), start=1):
        percent = 55 + int(idx / max(1, total_groups) * 35)
        _emit_progress(progress_callback, percent, f"Menganalisis kecocokan {group_name} ({idx}/{total_groups})...")

        matched_details = []
        missing_details = []

        for part in pcb_parts:
            key = _part_key(part)
            if key in group_parts:
                g_info = group_parts[key]
                matched_details.append({
                    "part_number": part,
                    "location_code": g_info["location_code"],
                    "type": g_info["type"],
                })
            else:
                missing_details.append(part)

        matched_count = len(matched_details)
        missing_count = len(missing_details)
        match_rate = round((matched_count / total_parts) * 100.0, 1)

        raw_matches.append({
            "group_name": group_name,
            "member_pcbs": group_pcbs_map.get(group_name, []),
            "total_pcb_parts": total_parts,
            "matched_count": matched_count,
            "missing_count": missing_count,
            "match_rate_percent": match_rate,
            "matched_details": matched_details,
            "missing_details": missing_details,
        })

    # Sort matches by match_rate desc, missing_count asc, group_name asc
    raw_matches.sort(key=lambda x: (-x["match_rate_percent"], x["missing_count"], x["group_name"]))

    group_matches = []
    for rank, item in enumerate(raw_matches, start=1):
        rate = item["match_rate_percent"]
        matched_cnt = item["matched_count"]
        missing_cnt = item["missing_count"]
        tot_cnt = item["total_pcb_parts"]

        if rank == 1:
            if rate >= 90.0:
                status = "SANGAT DIREKOMENDASIKAN (BEST MATCH)"
                note = f"Paling efisien! {matched_cnt} dari {tot_cnt} komponen ({rate}%) sudah terpasang. Hanya butuh {missing_cnt} feeder ekstra di slot kosong."
            elif rate >= 70.0:
                status = "DIREKOMENDASIKAN"
                note = f"Sangat cocok! {matched_cnt} dari {tot_cnt} komponen ({rate}%) ter-cover. Butuh {missing_cnt} feeder tambahan."
            else:
                status = "KAPASITAS TERBATAS"
                note = f"Kecocokan tertinggi saat ini {rate}%. Perlu {missing_cnt} feeder tambahan dipasang manual."
        else:
            if rate >= 85.0:
                status = "HIGH MATCH"
                note = f"Kecocokan tinggi ({rate}%). Butuh {missing_cnt} feeder tambahan."
            elif rate >= 65.0:
                status = "MODERATE MATCH"
                note = f"Kecocokan sedang ({rate}%). Butuh {missing_cnt} feeder tambahan."
            else:
                status = "LOW MATCH"
                note = f"Kecocokan rendah ({rate}%). {missing_cnt} komponen perlu pasang baru."

        group_matches.append(GroupMatchItem(
            rank=rank,
            group_name=item["group_name"],
            member_pcbs=item["member_pcbs"],
            total_pcb_parts=tot_cnt,
            matched_count=matched_cnt,
            missing_count=missing_cnt,
            match_rate_percent=rate,
            status_recommendation=status,
            recommendation_note=note,
            matched_details=item["matched_details"],
            missing_details=item["missing_details"],
        ))

    best_match = group_matches[0] if group_matches else None

    _emit_progress(progress_callback, 100, "Analisis kecocokan selesai!")
    return PcbGroupMatcherResult(
        output_path=config.output_path,
        pcb_name=pcb_name,
        import_mode=config.import_mode,
        pcb_parts_count=len(pcb_parts),
        unique_parts=pcb_parts,
        group_matches=group_matches,
        best_match=best_match,
    )


def export_pcb_group_matcher_result(result: PcbGroupMatcherResult, output_path: str):
    if not result or not result.group_matches:
        raise ServiceError("Tidak ada hasil analisis kecocokan untuk diexport.", title="Data kosong")

    out_path = Path(output_path)
    if out_path.suffix.lower() != ".xlsx":
        out_path = out_path.with_suffix(".xlsx")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    # Sheet 1: Ranking Summary
    ws_rank = wb.active
    ws_rank.title = "Group Match Ranking"

    ws_rank.append([
        "Rank",
        "Group Name",
        "Match Rate (%)",
        "Matched Parts",
        "Extra Feeders Needed",
        "Total PCB Parts",
        "Member PCBs in Group",
        "Status & Recommendation",
        "Note",
    ])

    for item in result.group_matches:
        ws_rank.append([
            item.rank,
            item.group_name,
            f"{item.match_rate_percent}%",
            f"{item.matched_count} / {item.total_pcb_parts}",
            item.missing_count,
            item.total_pcb_parts,
            ", ".join(item.member_pcbs) if item.member_pcbs else "-",
            item.status_recommendation,
            item.recommendation_note,
        ])

    _style_summary_sheet(ws_rank)

    # Sheet 2: Matched Components Detail
    ws_matched = wb.create_sheet("Matched Components")
    ws_matched.append(["Group Name", "Location Code", "Component Part Number", "Fix Type"])
    for item in result.group_matches:
        for m in item.matched_details:
            ws_matched.append([
                item.group_name,
                m["location_code"],
                m["part_number"],
                m["type"],
            ])
    _style_detail_sheet(ws_matched)

    # Sheet 3: Extra Components Needed (Change-Over)
    ws_extra = wb.create_sheet("Extra Components Needed")
    ws_extra.append(["Group Name", "Extra Component Part Number", "Action Needed"])
    for item in result.group_matches:
        for p in item.missing_details:
            ws_extra.append([
                item.group_name,
                p,
                "Butuh Feeder Tambahan (Pasang di Slot Kosong)",
            ])
    _style_detail_sheet(ws_extra)

    wb.save(out_path)
    return str(out_path)


def _extract_new_pcb_parts(mode, source_path):
    path = Path(source_path)
    parts_set = set()
    pcb_name = path.stem

    if mode == MODE_BOM_FILE:
        if path.suffix.lower() in EXCEL_EXTENSIONS:
            part_list = _read_parts_from_excel(path)
        elif path.suffix.lower() in TEXT_EXTENSIONS:
            part_list = _read_parts_from_text_bom(path)
        else:
            raise ServiceError(f"Ekstensi file BOM tidak didukung: {path.suffix}", title="Format file tidak didukung")
        for p in part_list:
            key = _part_key(p)
            if key:
                parts_set.add(p)

    elif mode == MODE_PROGRAM_EXCEL:
        if path.suffix.lower() not in EXCEL_EXTENSIONS:
            raise ServiceError("File program PCB harus berformat Excel (.xlsx, .xls, .xlsm).", title="Format file tidak valid")
        part_list = _read_parts_from_excel(path)
        for p in part_list:
            key = _part_key(p)
            if key:
                parts_set.add(p)

    elif mode == MODE_PROGRAM_FOLDER:
        if not path.is_dir():
            raise ServiceError(f"Folder PCB tidak ditemukan:\n{path}", title="Folder tidak ditemukan")
        pcb_name = path.name
        excel_files = scan_recursive_files(path, EXCEL_EXTENSIONS, skip_prefixes=("~$",))
        if not excel_files:
            raise ServiceError(f"Tidak ada file Excel program yang ditemukan dalam folder:\n{path}", title="File tidak ditemukan")

        for f in excel_files:
            try:
                part_list = _read_parts_from_excel(f)
                for p in part_list:
                    key = _part_key(p)
                    if key:
                        parts_set.add(p)
            except Exception:
                continue

    else:
        raise ServiceError(f"Mode import tidak dikenal: {mode}", title="Mode tidak valid")

    # Sort parts alphabetically
    sorted_parts = sorted(list(parts_set), key=lambda x: x.upper())
    return sorted_parts, pcb_name


def _read_parts_from_excel(path):
    try:
        part_values = _read_bom_parts(path)
        if part_values:
            return part_values
    except Exception:
        pass

    # Fallback openpyxl scan across worksheets
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        parts = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    val = _part_text(cell)
                    if val and not val.isdigit() and len(val) >= 4 and not val.upper().startswith(("PART", "SLOT", "LOCATION")):
                        parts.append(val)
        wb.close()
        return _clean_part_values(parts)
    except Exception:
        return []


def _read_parts_from_text_bom(path):
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
    except Exception as exc:
        raise ServiceError(f"Gagal membaca file text BOM:\n{path}", title="File Error") from exc

    lines = [line.strip() for line in content.splitlines() if line.strip()]
    if not lines:
        return []

    sample = lines[0]
    delim = "\t" if "\t" in sample else (";" if ";" in sample else ",")
    reader = csv.reader(lines, delimiter=delim)
    header = None
    part_col = -1
    raw_parts = []

    for row in reader:
        if not row:
            continue
        clean_row = [_part_text(c) for c in row]
        if header is None:
            upper_row = [c.upper() for c in clean_row]
            for idx, c in enumerate(upper_row):
                if "PART" in c or "P/N" in c or "COMPONENT" in c:
                    header = upper_row
                    part_col = idx
                    break
            if header is not None:
                continue

        if part_col >= 0 and len(clean_row) > part_col:
            val = clean_row[part_col]
        else:
            val = clean_row[0] if clean_row else ""

        if val and not val.upper().startswith(("PART", "P/N", "LOCATION", "SKIPPED")):
            raw_parts.append(val)

    return _clean_part_values(raw_parts)


def _load_fix_feeder_groups(fix_feeder_group_path):
    path = Path(fix_feeder_group_path)
    if not path.is_file():
        raise ServiceError(f"File Fix Feeder Group Excel tidak ditemukan:\n{path}", title="File tidak ditemukan")

    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        raise ServiceError("File Excel Fix Feeder Group tidak dapat dibaca.", title="Excel Error") from exc

    group_pcbs_map = defaultdict(list)
    if "Summary" in wb.sheetnames:
        ws_sum = wb["Summary"]
        header = None
        group_idx = -1
        pcb_idx = -1
        for row in ws_sum.iter_rows(values_only=True):
            vals = [_part_text(c) for c in row]
            if not any(vals):
                continue
            if header is None:
                header_upper = [v.upper() for v in vals]
                if "GROUP NAME" in header_upper and "PCB NAME" in header_upper:
                    header = header_upper
                    group_idx = header_upper.index("GROUP NAME")
                    pcb_idx = header_upper.index("PCB NAME")
            else:
                if group_idx >= 0 and pcb_idx >= 0 and len(vals) > max(group_idx, pcb_idx):
                    g_name = vals[group_idx]
                    p_name = vals[pcb_idx]
                    if g_name and p_name:
                        if p_name not in group_pcbs_map[g_name]:
                            group_pcbs_map[g_name].append(p_name)

    group_parts_map = OrderedDict()

    for worksheet in wb.worksheets:
        title = worksheet.title
        if title.lower() in {"summary", "summary sheet"}:
            continue

        parts_in_group = OrderedDict()
        header_map = {}
        for row_number, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
            raw_cells = [_part_text(c) for c in row]
            while raw_cells and not raw_cells[-1]:
                raw_cells.pop()
            if not any(raw_cells):
                continue

            cells_upper = [c.upper() for c in raw_cells]
            if "PART NUMBER" in cells_upper or "LOCATION CODE" in cells_upper or "TYPE" in cells_upper:
                for idx, c in enumerate(cells_upper):
                    if "PART NUMBER" in c or "PART" in c:
                        header_map["part"] = idx
                    elif "LOCATION" in c or "SLOT" in c:
                        header_map["loc"] = idx
                    elif "TYPE" in c:
                        header_map["type"] = idx
                continue

            part_num = ""
            loc_code = ""
            part_type = "FIXED"

            if "loc" in header_map and "part" in header_map:
                if len(raw_cells) > max(header_map["loc"], header_map["part"]):
                    loc_code = raw_cells[header_map["loc"]]
                    part_num = raw_cells[header_map["part"]]
                    if "type" in header_map and len(raw_cells) > header_map["type"]:
                        part_type = raw_cells[header_map["type"]].upper() or "FIXED"

            if not loc_code or not part_num:
                # Fallback check row for location code & part number
                for idx, c in enumerate(raw_cells):
                    if re.match(r"^\[(\d+)\](\d+)", c):
                        loc_code = c
                        break
                if loc_code:
                    for c in raw_cells:
                        if c != loc_code and c.upper() not in {"FIXED", "SUBSTITUTE", "DYNAMIC", "L", "R"} and not re.match(r"^\[(\d+)\](\d+)", c):
                            part_num = c
                            break

            if loc_code and part_num and part_type in {"FIXED", "SUBSTITUTE"}:
                key = _part_key(part_num)
                if key:
                    parts_in_group[key] = {
                        "part_number": part_num,
                        "location_code": loc_code,
                        "type": part_type,
                    }

        if parts_in_group:
            group_parts_map[title] = parts_in_group

    wb.close()
    return group_parts_map, group_pcbs_map


def _validate_config(config: PcbGroupMatcherConfig):
    if not config.import_mode:
        raise ServiceError("Metode import PCB belum dipilih.", title="Input belum lengkap")
    if not config.pcb_source_path:
        raise ServiceError("File / Folder PCB baru belum dipilih.", title="Input belum lengkap")
    if not config.fix_feeder_group_path:
        raise ServiceError("File Fix Feeder Group Excel belum dipilih.", title="Input belum lengkap")


def _emit_progress(progress_callback, percent, message):
    if progress_callback:
        progress_callback(percent, message)


def _style_summary_sheet(ws):
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            if col_idx in {1, 3, 4, 5, 6}:
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    ws.freeze_panes = "A2"
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)


def _style_detail_sheet(ws):
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    left_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            cell.alignment = left_align

    ws.freeze_panes = "A2"
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 15)
