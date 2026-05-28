import re
import shlex
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from services.errors import ServiceError
from utils.encoding import read_lines_with_fallback


OUTPUT_HEADERS = ["Table", "Slot", "Position", "Location Code", "Part Number"]


@dataclass
class FeederMappingResult:
    records: list
    source_file: str
    row_count: int
    table_count: int
    part_count: int


def suggest_output_name(source_path):
    stem = _clean_filename_part(Path(source_path or "Feeder_Mapping").stem) or "Feeder_Mapping"
    return f"Feeder_Mapping_Result_{stem}.xlsx"


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


def export_feeder_mapping(records, output_path):
    output = Path(output_path)
    if output.suffix.lower() != ".xlsx":
        output = output.with_suffix(".xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Detailed Feeder Setup"
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

    _style_workbook(worksheet)
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


def _style_workbook(worksheet):
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


def _clean_filename_part(value):
    text = str(value or "").strip()
    text = re.sub(r'[<>:"/\\|?*]+', "_", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" .")


def _clean_path(file_path):
    return str(file_path or "").strip().replace('"', "").replace("'", "")
