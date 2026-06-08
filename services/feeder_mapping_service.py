import re
import shlex
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from services.errors import ServiceError
from utils.encoding import read_lines_with_fallback
from utils.sort import natural_sort_key


OUTPUT_HEADERS = ["Table", "Slot", "Position", "Location Code", "Part Number"]
SUMMARY_HEADERS = [
    "Part Number",
    "Most Common Location",
    "Location Hit Count",
    "Files With Part",
    "Total Files",
    "Coverage %",
    "Total Feeders",
    "Avg Feeders/Selected File",
    "Location Summary",
    "Feeder Files",
]


@dataclass
class FeederMappingResult:
    records: list
    source_file: str
    row_count: int
    table_count: int
    part_count: int


@dataclass
class FeederMappingBatchResult:
    mappings: list
    summary_records: list
    source_count: int
    row_count: int
    part_count: int
    output_path: str = ""


def suggest_output_name(source_path):
    stem = _clean_filename_part(Path(source_path or "Feeder_Mapping").stem) or "Feeder_Mapping"
    return f"Feeder_Mapping_Result_{stem}.xlsx"


def suggest_multiple_output_name(source_paths=None):
    paths = list(source_paths or [])
    if len(paths) == 1:
        stem = _clean_filename_part(Path(paths[0]).stem) or "Multiple"
        return f"Feeder_Mapping_Multiple_{stem}.xlsx"
    return "Feeder_Mapping_Multiple_Result.xlsx"


def load_feeder_mapping(file_path):
    path = _clean_path(file_path)
    if not Path(path).is_file():
        raise ServiceError(f"File mesin NPM tidak ditemukan:\n{path}", title="File tidak ditemukan")

    try:
        lines, _ = read_lines_with_fallback(path)
    except Exception as exc:
        raise ServiceError("File mesin NPM tidak bisa dibaca.", title="Encoding error") from exc

    parts_rows = _read_first_available_section_rows(lines, ("PartsData", "PartsDataEx"))
    feeder_rows = _read_section_rows(lines, "FeederData")
    fixed_rows = _read_section_rows(lines, "FixedFeeder")

    part_lookup = _build_lookup(parts_rows, "IDNUM")
    feeder_lookup = _build_lookup(feeder_rows, "IDNUM")
    records = _build_mapping_records(fixed_rows, part_lookup, feeder_lookup)

    if not records:
        raise ServiceError("Tidak ada Fixed Feeder aktif yang bisa dikonversi.", title="Data kosong")

    table_count = len({record["table"] for record in records})
    part_count = len({record["part_number"] for record in records})
    return FeederMappingResult(
        records=records,
        source_file=Path(path).name,
        row_count=len(records),
        table_count=table_count,
        part_count=part_count,
    )


def generate_feeder_mapping_excel(source_path, output_path):
    result = load_feeder_mapping(source_path)
    exported_path = export_feeder_mapping(result.records, output_path)
    return result, exported_path


def load_multiple_feeder_mappings(source_paths):
    paths = _clean_source_paths(source_paths)
    if not paths:
        raise ServiceError("Belum ada file export mesin NPM yang dipilih.", title="Input belum lengkap")

    mappings = [load_feeder_mapping(path) for path in paths]
    summary_records = _build_summary_records(mappings)
    return FeederMappingBatchResult(
        mappings=mappings,
        summary_records=summary_records,
        source_count=len(mappings),
        row_count=sum(mapping.row_count for mapping in mappings),
        part_count=len(summary_records),
    )


def generate_multiple_feeder_mapping_excel(source_paths, output_path):
    result = load_multiple_feeder_mappings(source_paths)
    result.output_path = export_multiple_feeder_mapping(result, output_path)
    return result


def export_feeder_mapping(records, output_path):
    output = Path(output_path)
    if output.suffix.lower() != ".xlsx":
        output = output.with_suffix(".xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Detailed Feeder Setup"
    _append_mapping_sheet(worksheet, records)
    workbook.save(output)
    return str(output)


def export_multiple_feeder_mapping(result, output_path):
    output = Path(output_path)
    if output.suffix.lower() != ".xlsx":
        output = output.with_suffix(".xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "Summary"
    _append_summary_sheet(summary_sheet, result)

    used_titles = {"summary"}
    for mapping in result.mappings:
        worksheet = workbook.create_sheet(_unique_sheet_title(mapping.source_file, used_titles))
        _append_mapping_sheet(worksheet, mapping.records)

    workbook.save(output)
    return str(output)


def _build_mapping_records(fixed_rows, part_lookup, feeder_lookup):
    records = []
    for row in fixed_rows:
        table, slot = _decode_pickup_unit(row.get("PU", ""))
        if not table or not slot:
            continue

        feeder_a = row.get("FeederA", "")
        part_a = row.get("PartsA", "")
        if not _is_active(feeder_a, part_a, part_lookup):
            continue

        kind = _feeder_kind(feeder_lookup.get(feeder_a, {}))
        if kind == 2:
            records.append(_record(table, slot, "L", f"[{table}]{slot}L", part_lookup[part_a]["NAME"]))
            feeder_b = row.get("FeederB", "")
            part_b = row.get("PartsB", "")
            if _is_active(feeder_b, part_b, part_lookup):
                records.append(_record(table, slot, "R", f"[{table}]{slot}R", part_lookup[part_b]["NAME"]))
            continue

        if kind == 3:
            records.append(_record(table, slot, "Large (2-Rel)", f"[{table}]{slot}", part_lookup[part_a]["NAME"]))
            continue

        if kind == 4:
            records.append(
                _record(
                    table,
                    slot,
                    "Extra Large (3-Rel)",
                    f"[{table}]{slot}-{slot + 1}",
                    part_lookup[part_a]["NAME"],
                )
            )
            continue

        records.append(_record(table, slot, f"Kind {kind or 'Unknown'}", f"[{table}]{slot}", part_lookup[part_a]["NAME"]))

    position_order = {"L": 0, "R": 1, "Large (2-Rel)": 0, "Extra Large (3-Rel)": 0}
    records.sort(key=lambda item: (item["table"], item["slot"], position_order.get(item["position"], 9)))
    return records


def _read_section_rows(lines, section_name):
    section_lines = _extract_section(lines, section_name)
    if len(section_lines) < 2:
        raise ServiceError(f"Section [{section_name}] kosong atau tidak punya header.", title="Format NPM tidak valid")

    header_line_number, header_line = section_lines[1]
    header = _split_line(header_line, section_name, header_line_number)
    rows = []
    for line_number, line in section_lines[2:]:
        if not line.strip():
            continue
        values = _split_line(line, section_name, line_number)
        if len(values) != len(header):
            raise ServiceError(
                f"Jumlah kolom tidak sesuai di [{section_name}] line {line_number}.\n"
                f"Expected {len(header)}, got {len(values)}.",
                title="Format NPM tidak valid",
            )
        rows.append(dict(zip(header, values)))
    return rows


def _read_first_available_section_rows(lines, section_names):
    for section_name in section_names:
        if _has_section(lines, section_name):
            return _read_section_rows(lines, section_name)

    names = ", ".join(f"[{section_name}]" for section_name in section_names)
    raise ServiceError(
        f"Section {names} tidak ditemukan.\n"
        "File export NPM harus berisi data part number.",
        title="Format NPM tidak valid",
    )


def _extract_section(lines, section_name):
    start_index = None
    section_marker = f"[{section_name}]"
    for index, line in enumerate(lines):
        if line.strip().lower() == section_marker.lower():
            start_index = index
            break

    if start_index is None:
        raise ServiceError(f"Section [{section_name}] tidak ditemukan.", title="Format NPM tidak valid")

    end_index = len(lines)
    for index in range(start_index + 1, len(lines)):
        stripped = lines[index].strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            end_index = index
            break

    return [(line_number, lines[line_number - 1].strip()) for line_number in range(start_index + 1, end_index + 1)]


def _has_section(lines, section_name):
    section_marker = f"[{section_name}]"
    return any(line.strip().lower() == section_marker.lower() for line in lines)


def _split_line(line, section_name, line_number):
    try:
        return shlex.split(line, posix=True)
    except ValueError as exc:
        raise ServiceError(
            f"Gagal membaca [{section_name}] line {line_number}:\n{exc}",
            title="Format NPM tidak valid",
        ) from exc


def _build_lookup(rows, key):
    lookup = {}
    for row in rows:
        value = str(row.get(key, "")).strip()
        if value:
            lookup[value] = row
    return lookup


def _decode_pickup_unit(value):
    try:
        pickup_unit = int(float(str(value).strip()))
    except (TypeError, ValueError):
        return 0, 0
    if pickup_unit <= 0:
        return 0, 0
    return pickup_unit // 10000, pickup_unit % 10000


def _is_active(feeder_id, part_id, part_lookup):
    return (
        str(feeder_id).strip() not in {"", "0", "-1"}
        and str(part_id).strip() not in {"", "0", "-1"}
        and str(part_id).strip() in part_lookup
    )


def _feeder_kind(feeder_row):
    try:
        return int(float(str(feeder_row.get("Kind", "0")).strip()))
    except (TypeError, ValueError):
        return 0


def _record(table, slot, position, location_code, part_number):
    return {
        "table": table,
        "slot": slot,
        "position": position,
        "location_code": location_code,
        "part_number": part_number,
    }


def _append_mapping_sheet(worksheet, records):
    worksheet.append(OUTPUT_HEADERS)

    for record in records:
        worksheet.append(
            [
                record["table"],
                record["slot"],
                record["position"],
                record["location_code"],
                record["part_number"],
            ]
        )

    _style_mapping_sheet(worksheet)


def _append_summary_sheet(worksheet, result):
    worksheet.append(["Metric", "Value"])
    worksheet.append(["Feeder File Count", result.source_count])
    worksheet.append(["Total Feeder Rows", result.row_count])
    worksheet.append(["Unique Parts", result.part_count])
    worksheet.append([])
    header_row = worksheet.max_row + 1
    worksheet.append(SUMMARY_HEADERS)

    for record in result.summary_records:
        worksheet.append(
            [
                record["part_number"],
                record["most_common_location"],
                record["location_hit_count"],
                record["files_with_part"],
                record["total_files"],
                record["coverage_percent"],
                record["total_feeders"],
                record["avg_feeders_per_selected_file"],
                record["location_summary"],
                record["feeder_files"],
            ]
        )

    _style_summary_sheet(worksheet, header_row)


def _build_summary_records(mappings):
    total_files = len(mappings)
    stats = {}

    for mapping in mappings:
        per_file = {}
        for record in mapping.records:
            part_number = str(record.get("part_number", "")).strip()
            if not part_number:
                continue
            key = part_number.upper()
            part_data = per_file.setdefault(
                key,
                {
                    "part_number": part_number,
                    "location_codes": [],
                },
            )
            part_data["location_codes"].append(str(record.get("location_code", "")).strip())

        for key, part_data in per_file.items():
            item = stats.setdefault(
                key,
                {
                    "part_number": part_data["part_number"],
                    "files": [],
                    "total_feeders": 0,
                    "location_counts": Counter(),
                },
            )
            feeder_count = len(part_data["location_codes"])
            item["files"].append(mapping.source_file)
            item["total_feeders"] += feeder_count
            item["location_counts"].update(location for location in part_data["location_codes"] if location)

    summary_records = []
    for item in stats.values():
        files_with_part = len(item["files"])
        total_feeders = item["total_feeders"]
        location_items = _sorted_location_counts(item["location_counts"])
        top_count = location_items[0][1] if location_items else 0
        most_common_locations = [location for location, count in location_items if count == top_count]
        summary_records.append(
            {
                "part_number": item["part_number"],
                "most_common_location": ", ".join(most_common_locations),
                "location_hit_count": top_count,
                "files_with_part": files_with_part,
                "total_files": total_files,
                "coverage_percent": round((files_with_part / total_files) * 100, 2) if total_files else 0,
                "total_feeders": total_feeders,
                "avg_feeders_per_selected_file": round(total_feeders / total_files, 2) if total_files else 0,
                "location_summary": ", ".join(f"{location} ({count})" for location, count in location_items),
                "feeder_files": ", ".join(item["files"]),
            }
        )

    summary_records.sort(
        key=lambda row: (
            -row["files_with_part"],
            natural_sort_key(row["part_number"]),
        )
    )
    return summary_records


def _sorted_location_counts(location_counts):
    return sorted(
        location_counts.items(),
        key=lambda item: (
            -item[1],
            natural_sort_key(item[0]),
        ),
    )


def _style_mapping_sheet(worksheet):
    header_fill = PatternFill("solid", fgColor="FF2B3A4C")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    default_font = Font(name="Calibri", size=11)
    border_side = Side(style="thin", color="FFD3D3D3")
    border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    widths = {"A": 9, "B": 8, "C": 23, "D": 17, "E": 15}
    for column_letter, width in widths.items():
        worksheet.column_dimensions[column_letter].width = width

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = default_font
            cell.border = border
            horizontal = "left" if cell.column in (4, 5) else "center"
            cell.alignment = Alignment(horizontal=horizontal, vertical="center")

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions


def _style_summary_sheet(worksheet, header_row):
    header_fill = PatternFill("solid", fgColor="FF2B3A4C")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    default_font = Font(name="Calibri", size=11)
    metric_font = Font(name="Calibri", size=11, bold=True)
    border_side = Side(style="thin", color="FFD3D3D3")
    border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    widths = {
        "A": 28,
        "B": 22,
        "C": 16,
        "D": 16,
        "E": 13,
        "F": 12,
        "G": 14,
        "H": 26,
        "I": 44,
        "J": 70,
    }
    for column_letter, width in widths.items():
        worksheet.column_dimensions[column_letter].width = width

    for row in worksheet.iter_rows(min_row=1, max_row=4, max_col=2):
        row[0].font = metric_font
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="center")

    for cell in worksheet[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row in worksheet.iter_rows(min_row=header_row + 1):
        for cell in row:
            cell.font = default_font
            cell.border = border
            horizontal = "left" if cell.column in (1, 2, 9, 10) else "center"
            cell.alignment = Alignment(horizontal=horizontal, vertical="center")

    worksheet.freeze_panes = f"A{header_row + 1}"
    worksheet.auto_filter.ref = f"A{header_row}:J{worksheet.max_row}"


def _unique_sheet_title(source_file, used_titles):
    base = _clean_sheet_title(Path(source_file or "Feeder Setup").stem) or "Feeder Setup"
    if base.lower() == "summary":
        base = "Summary Detail"

    title = base[:31]
    counter = 2
    while title.lower() in used_titles:
        suffix = f" ({counter})"
        title = f"{base[:31 - len(suffix)]}{suffix}"
        counter += 1

    used_titles.add(title.lower())
    return title


def _clean_sheet_title(value):
    text = str(value or "").strip()
    text = re.sub(r"[\[\]:*?/\\]+", "_", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" .'")


def _clean_source_paths(source_paths):
    if isinstance(source_paths, (str, Path)):
        raw_paths = [source_paths]
    else:
        raw_paths = list(source_paths or [])

    paths = []
    seen = set()
    for raw_path in raw_paths:
        path = _clean_path(raw_path)
        if not path:
            continue
        normalized = str(Path(path))
        key = normalized.lower()
        if key in seen:
            continue
        seen.add(key)
        paths.append(normalized)
    return paths


def _clean_filename_part(value):
    text = str(value or "").strip()
    text = re.sub(r'[<>:"/\\|?*]+', "_", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" .")


def _clean_path(file_path):
    return str(file_path or "").strip().replace('"', "").replace("'", "")
