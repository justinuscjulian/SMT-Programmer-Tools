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
    "Feeder Paling Sering",
    "Jumlah File di Feeder Itu",
    "File Component Muncul",
    "Total File Mapping",
    "Total Muncul",
    "Feeder Lain",
]
DEFAULT_BALANCING_PART_NUMBERS = [
    "0RJ1002C678",
    "0CK104BF56A",
    "0RJ0000C678",
    "0CK104BH56A",
    "0RJ1000C678",
    "EAE52158501",
    "EAE66302101",
]
FIXED_FEEDER_SIDES = "ABCDEFGHIJ"
VALID_NPM_EXPORT_SUFFIXES = {".txt", ".crb"}
CM602_FEEDER_SIDES = "ABCDEFGHIJ"


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


@dataclass
class NpmFeederImportResult:
    output_path: str
    target_file: str
    reference_count: int
    target_part_count: int
    assigned_part_count: int
    assignment_count: int
    balancing_part_count: int
    missing_recommendation_parts: list
    missing_location_rows: list
    conflict_rows: list


def suggest_output_name(source_path):
    stem = _clean_filename_part(Path(source_path or "Feeder_Mapping").stem) or "Feeder_Mapping"
    return f"Feeder_Mapping_Result_{stem}.xlsx"


def suggest_multiple_output_name(source_paths=None):
    paths = list(source_paths or [])
    if len(paths) == 1:
        stem = _clean_filename_part(Path(paths[0]).stem) or "Multiple"
        return f"Feeder_Mapping_Multiple_{stem}.xlsx"
    return "Feeder_Mapping_Multiple_Result.xlsx"


def suggest_cm602_output_name(source_path):
    stem = _clean_filename_part(Path(source_path or "CM602_Feeder_Mapping").stem) or "CM602_Feeder_Mapping"
    return f"CM602_Feeder_Mapping_Result_{stem}.xlsx"


def suggest_npm_import_output_name(target_path):
    stem = _clean_filename_part(Path(target_path or "NPM_Program").stem) or "NPM_Program"
    return f"{stem}_FEEDER_SETUP.crb"


def default_balancing_part_numbers_text():
    return "\n".join(DEFAULT_BALANCING_PART_NUMBERS)


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


def load_cm602_feeder_mapping(file_path):
    path = _clean_path(file_path)
    if not Path(path).is_file():
        raise ServiceError(f"File CM602 tidak ditemukan:\n{path}", title="File tidak ditemukan")

    try:
        lines, _ = read_lines_with_fallback(path)
    except Exception as exc:
        raise ServiceError("File CM602 tidak bisa dibaca.", title="Encoding error") from exc

    if _has_section(lines, "FeederFix"):
        feeder_fix_rows = _read_section_rows(lines, "FeederFix", machine_label="CM602")
        records = _build_cm602_feeder_fix_records(feeder_fix_rows)
    elif _has_section(lines, "StockData") and _has_section(lines, "PartsData"):
        stock_rows = _read_section_rows(lines, "StockData", machine_label="CM602")
        part_lookup = _build_lookup(_read_section_rows(lines, "PartsData", machine_label="CM602"), "IDNUM")
        records = _build_cm602_stock_records(stock_rows, part_lookup)
    else:
        raise ServiceError(
            "File CM602 harus berisi section [FeederFix] atau kombinasi [StockData] dan [PartsData].",
            title="Format CM602 tidak valid",
        )

    if not records:
        raise ServiceError("Tidak ada feeder CM602 aktif yang bisa dikonversi.", title="Data kosong")

    table_count = len({record["table"] for record in records})
    part_count = len({record["part_number"] for record in records})
    return FeederMappingResult(
        records=records,
        source_file=Path(path).name,
        row_count=len(records),
        table_count=table_count,
        part_count=part_count,
    )


def generate_cm602_feeder_mapping_excel(source_path, output_path):
    result = load_cm602_feeder_mapping(source_path)
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


def generate_npm_feeder_import_file(target_path, reference_folder, output_path, balancing_part_numbers=None):
    target = Path(_clean_path(target_path))
    if not target.is_file():
        raise ServiceError(f"Target NPM program tidak ditemukan:\n{target}", title="File tidak ditemukan")

    reference_paths = _collect_reference_files(reference_folder, exclude_path=target)
    if not reference_paths:
        raise ServiceError("Folder reference tidak berisi file .txt/.crb yang valid.", title="Reference kosong")

    output = Path(output_path)
    if output.suffix.lower() != ".crb":
        output = output.with_suffix(".crb")
    output.parent.mkdir(parents=True, exist_ok=True)

    target_lines, encoding = read_lines_with_fallback(target)
    reference_mappings = [load_feeder_mapping(path) for path in reference_paths]
    balancing_parts = _parse_balancing_part_numbers(balancing_part_numbers)
    recommendations = _build_npm_feeder_recommendations(reference_mappings, balancing_parts)

    fixed_rows = _read_section_rows(target_lines, "FixedFeeder")
    part_lookup = _build_lookup(_read_section_rows(target_lines, "PartsData"), "IDNUM")
    feeder_lookup = _build_lookup(_read_section_rows(target_lines, "FeederData"), "IDNUM")
    target_part_ids = _target_used_part_ids(target_lines, part_lookup)
    if not target_part_ids:
        raise ServiceError("Target NPM program tidak punya part placement aktif.", title="Data kosong")
    target_parts = _unique_target_parts(target_part_ids, part_lookup)

    fixed_header = _fixed_feeder_header(target_lines)
    fixed_rows_by_pu = {str(row.get("PU", "")).strip(): dict(row) for row in fixed_rows}
    existing_assignments = _existing_fixed_assignments(fixed_rows, part_lookup, feeder_lookup)
    new_fixed_rows = [_empty_fixed_row(row) for row in fixed_rows]
    new_fixed_by_pu = {str(row.get("PU", "")).strip(): row for row in new_fixed_rows}
    occupied = {}

    assigned_part_keys = set()
    assignment_count = 0
    missing_recommendation_parts = []
    missing_location_rows = []
    conflict_rows = []
    balancing_part_keys = set()

    for part_key, target_part in sorted(target_parts.items(), key=lambda item: natural_sort_key(item[1]["part_number"])):
        part_id = target_part["part_id"]
        part_row = target_part["part_row"]
        part_number = target_part["part_number"]
        recommendation = recommendations.get(part_key)
        assignments = []

        if recommendation:
            assignments = [
                {
                    "location_code": location_code,
                    "feeder_id": recommendation["feeder_ids"].get(location_code, ""),
                }
                for location_code in recommendation["locations"]
            ]
        else:
            assignments = list(existing_assignments.get(part_key, []))
            if not assignments:
                missing_recommendation_parts.append(part_number)
                continue

        assigned_this_part = False
        if len(assignments) > 1:
            balancing_part_keys.add(part_key)

        for assignment in assignments:
            location_code = assignment.get("location_code", "")
            parsed_location = _parse_location_code(location_code)
            if not parsed_location:
                missing_location_rows.append(f"{part_number}: {location_code}")
                continue
            pu, side = parsed_location
            row = new_fixed_by_pu.get(str(pu))
            if row is None:
                missing_location_rows.append(f"{part_number}: {location_code}")
                continue

            location_key = (str(pu), side)
            current_part = occupied.get(location_key)
            if current_part and current_part != part_number:
                conflict_rows.append(f"{location_code}: {current_part} vs {part_number}")
                continue

            feeder_id = assignment.get("feeder_id") or ""
            if feeder_id not in feeder_lookup:
                feeder_id = _existing_feeder_id_for_location(existing_assignments, part_key, location_code)
            if feeder_id not in feeder_lookup:
                feeder_id = _default_feeder_id(part_row, feeder_lookup)
            if not feeder_id:
                missing_location_rows.append(f"{part_number}: feeder type tidak ditemukan")
                continue

            _set_fixed_assignment(row, side, feeder_id, part_id)
            occupied[location_key] = part_number
            assigned_this_part = True
            assignment_count += 1

        if assigned_this_part:
            assigned_part_keys.add(part_key)

    output_lines = _replace_fixed_feeder_section(target_lines, fixed_header, new_fixed_rows)
    with output.open("w", encoding=encoding, newline="") as handle:
        handle.writelines(output_lines)

    return NpmFeederImportResult(
        output_path=str(output),
        target_file=target.name,
        reference_count=len(reference_paths),
        target_part_count=len(target_parts),
        assigned_part_count=len(assigned_part_keys),
        assignment_count=assignment_count,
        balancing_part_count=len(balancing_part_keys),
        missing_recommendation_parts=missing_recommendation_parts,
        missing_location_rows=missing_location_rows,
        conflict_rows=conflict_rows,
    )


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
        feeder_b = row.get("FeederB", "")
        part_a = row.get("PartsA", "")
        part_b = row.get("PartsB", "")
        active_a = _is_active(feeder_a, part_a, part_lookup)
        active_b = _is_active(feeder_b, part_b, part_lookup)
        if not active_a and not active_b:
            continue

        feeder_for_kind = feeder_a if active_a else feeder_b
        kind = _feeder_kind(feeder_lookup.get(feeder_for_kind, {}))
        if kind == 2:
            if active_a:
                records.append(_record(table, slot, "L", f"[{table}]{slot}L", part_lookup[part_a]["NAME"], feeder_a))
            if active_b:
                records.append(_record(table, slot, "R", f"[{table}]{slot}R", part_lookup[part_b]["NAME"], feeder_b))
            continue

        if not active_a:
            records.append(_record(table, slot, f"Kind {kind or 'Unknown'} R", f"[{table}]{slot}R", part_lookup[part_b]["NAME"], feeder_b))
            continue

        if kind == 3:
            records.append(_record(table, slot, "Large (2-Rel)", f"[{table}]{slot}", part_lookup[part_a]["NAME"], feeder_a))
            continue

        if kind == 4:
            records.append(
                _record(
                    table,
                    slot,
                    "Extra Large (3-Rel)",
                    f"[{table}]{slot}-{slot + 1}",
                    part_lookup[part_a]["NAME"],
                    feeder_a,
                )
            )
            continue

        records.append(_record(table, slot, f"Kind {kind or 'Unknown'}", f"[{table}]{slot}", part_lookup[part_a]["NAME"], feeder_a))

    position_order = {"L": 0, "R": 1, "Large (2-Rel)": 0, "Extra Large (3-Rel)": 0}
    records.sort(key=lambda item: (item["table"], item["slot"], position_order.get(item["position"], 9)))
    return records


def _build_cm602_feeder_fix_records(feeder_fix_rows):
    records = []
    for row in feeder_fix_rows:
        table = _safe_int(row.get("BeamNo", ""))
        slot = _safe_int(row.get("PU", ""))
        if not table or not slot:
            continue

        for side in CM602_FEEDER_SIDES:
            part_number = _clean_cm602_part(row.get(f"Parts{side}", ""))
            feeder_id = str(row.get(f"Feeder{side}", "")).strip()
            if not _is_cm602_assignment_active(feeder_id, part_number):
                continue

            position = _cm602_position_label(side)
            records.append(_record(table, slot, position, f"[{table}]{slot}{position}", part_number, feeder_id))

    records.sort(key=_cm602_record_sort_key)
    return records


def _build_cm602_stock_records(stock_rows, part_lookup):
    records = []
    for row in stock_rows:
        table, slot = _decode_pickup_unit(row.get("N", ""))
        if not table or not slot:
            continue

        for side in CM602_FEEDER_SIDES:
            part_id = str(row.get(f"P{side}", "")).strip()
            feeder_id = str(row.get(f"T{side}", "")).strip()
            if part_id in {"", "0", "-1"} or part_id not in part_lookup:
                continue
            part_number = _clean_cm602_part(part_lookup[part_id].get("NAME", ""))
            if not _is_cm602_assignment_active(feeder_id, part_number):
                continue

            position = _cm602_position_label(side)
            records.append(_record(table, slot, position, f"[{table}]{slot}{position}", part_number, feeder_id))

    records.sort(key=_cm602_record_sort_key)
    return records


def _read_section_rows(lines, section_name, machine_label="NPM"):
    section_lines = _extract_section(lines, section_name)
    if len(section_lines) < 2:
        raise ServiceError(f"Section [{section_name}] kosong atau tidak punya header.", title=f"Format {machine_label} tidak valid")

    header_line_number, header_line = section_lines[1]
    header = _split_line(header_line, section_name, header_line_number, machine_label)
    rows = []
    for line_number, line in section_lines[2:]:
        if not line.strip():
            continue
        values = _split_line(line, section_name, line_number, machine_label)
        if len(values) != len(header):
            raise ServiceError(
                f"Jumlah kolom tidak sesuai di [{section_name}] line {line_number}.\n"
                f"Expected {len(header)}, got {len(values)}.",
                title=f"Format {machine_label} tidak valid",
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


def _split_line(line, section_name, line_number, machine_label="NPM"):
    try:
        return shlex.split(line, posix=True)
    except ValueError as exc:
        raise ServiceError(
            f"Gagal membaca [{section_name}] line {line_number}:\n{exc}",
            title=f"Format {machine_label} tidak valid",
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


def _safe_int(value):
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return 0


def _clean_cm602_part(value):
    return str(value or "").strip().strip('"').strip()


def _is_cm602_assignment_active(feeder_id, part_number):
    return (
        str(feeder_id).strip() not in {"", "0", "-1"}
        and str(part_number).strip() not in {"", "0", "-1"}
    )


def _cm602_position_label(side):
    side = str(side or "").upper()
    if side == "A":
        return "L"
    if side == "B":
        return "R"
    return side


def _cm602_record_sort_key(record):
    position_order = {"L": 0, "R": 1}
    position = str(record.get("position", ""))
    return (
        record.get("table", 0),
        record.get("slot", 0),
        position_order.get(position, 2 + CM602_FEEDER_SIDES.find(position) if position in CM602_FEEDER_SIDES else 99),
    )


def _record(table, slot, position, location_code, part_number, feeder_id=""):
    return {
        "table": table,
        "slot": slot,
        "position": position,
        "location_code": location_code,
        "part_number": part_number,
        "feeder_id": feeder_id,
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
    header_row = 1
    worksheet.append(SUMMARY_HEADERS)

    for record in result.summary_records:
        worksheet.append(
            [
                record["part_number"],
                record["top_location"],
                record["top_count"],
                record["file_count"],
                result.source_count,
                record["total_count"],
                record["other_locations"],
            ]
        )

    _style_summary_sheet(worksheet, header_row)


def _build_summary_records(mappings):
    stats = {}

    for mapping in mappings:
        per_file_locations = {}
        for record in mapping.records:
            part_number = str(record.get("part_number", "")).strip()
            if not part_number:
                continue
            key = part_number.upper()
            item = stats.setdefault(
                key,
                {
                    "part_number": part_number,
                    "location_counts": Counter(),
                    "location_file_counts": Counter(),
                    "file_count": 0,
                },
            )
            location_code = str(record.get("location_code", "")).strip()
            if location_code:
                item["location_counts"][location_code] += 1
                per_file_locations.setdefault(key, set()).add(location_code)

        for key, locations in per_file_locations.items():
            item = stats[key]
            item["file_count"] += 1
            for location_code in locations:
                item["location_file_counts"][location_code] += 1

    summary_records = []
    for item in stats.values():
        location_items = _sorted_location_counts(item["location_file_counts"])
        top_count = location_items[0][1] if location_items else 0
        top_locations = [location for location, count in location_items if count == top_count]
        other_locations = [(location, count) for location, count in location_items if count != top_count]
        summary_records.append(
            {
                "part_number": item["part_number"],
                "top_location": ", ".join(top_locations),
                "top_count": top_count,
                "file_count": item["file_count"],
                "total_count": sum(item["location_counts"].values()),
                "other_locations": ", ".join(f"{location} ({count} file)" for location, count in other_locations),
            }
        )

    summary_records.sort(
        key=lambda row: (
            -row["top_count"],
            -row["file_count"],
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


def _collect_reference_files(reference_folder, exclude_path=None):
    folder = Path(_clean_path(reference_folder))
    if not folder.is_dir():
        raise ServiceError(f"Folder reference tidak ditemukan:\n{folder}", title="Folder tidak ditemukan")

    exclude_key = str(Path(exclude_path).resolve()).lower() if exclude_path else ""
    paths = []
    for path in folder.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in VALID_NPM_EXPORT_SUFFIXES:
            continue
        if str(path.resolve()).lower() == exclude_key:
            continue
        paths.append(str(path))
    paths.sort(key=lambda value: natural_sort_key(Path(value).name))
    return paths


def _parse_balancing_part_numbers(value):
    if value is None:
        values = DEFAULT_BALANCING_PART_NUMBERS
    elif isinstance(value, str):
        values = re.split(r"[\s,;]+", value)
    else:
        values = value
    return {_part_key(item) for item in values if _part_key(item)}


def _build_npm_feeder_recommendations(mappings, balancing_parts):
    stats = {}
    for mapping in mappings:
        per_file_locations = {}
        for record in mapping.records:
            part_number = str(record.get("part_number", "")).strip()
            location_code = str(record.get("location_code", "")).strip()
            if not part_number or not location_code:
                continue

            key = _part_key(part_number)
            item = stats.setdefault(
                key,
                {
                    "part_number": part_number,
                    "location_counts": Counter(),
                    "feeder_counts": {},
                    "set_counts": Counter(),
                    "file_count": 0,
                    "multi_file_count": 0,
                },
            )
            item["location_counts"][location_code] += 1
            item["feeder_counts"].setdefault(location_code, Counter())[str(record.get("feeder_id", "")).strip()] += 1
            per_file_locations.setdefault(key, set()).add(location_code)

        for key, locations in per_file_locations.items():
            item = stats[key]
            location_set = tuple(sorted(locations, key=natural_sort_key))
            item["set_counts"][location_set] += 1
            item["file_count"] += 1
            if len(location_set) > 1:
                item["multi_file_count"] += 1

    recommendations = {}
    for key, item in stats.items():
        locations = _recommended_locations_for_part(key, item, balancing_parts)
        if not locations:
            continue
        feeder_ids = {}
        for location_code in locations:
            feeder_counts = item["feeder_counts"].get(location_code, Counter())
            feeder_ids[location_code] = _most_common_non_empty(feeder_counts)
        recommendations[key] = {
            "part_number": item["part_number"],
            "locations": locations,
            "feeder_ids": feeder_ids,
        }
    return recommendations


def _recommended_locations_for_part(part_key, item, balancing_parts):
    location_items = _sorted_location_counts(item["location_counts"])
    if not location_items:
        return []

    if part_key in balancing_parts:
        top_count = location_items[0][1]
        min_count = max(2, int(top_count * 0.25))
        return [location for location, count in location_items if count >= min_count]

    set_items = item["set_counts"].most_common()
    if set_items:
        top_set, top_set_count = set_items[0]
        enough_multi_files = item["multi_file_count"] >= max(3, int(item["file_count"] * 0.5))
        if len(top_set) > 1 and enough_multi_files and top_set_count >= 3:
            return list(top_set)

    return [location_items[0][0]]


def _most_common_non_empty(counter):
    for value, _ in counter.most_common():
        if str(value).strip() not in {"", "0", "-1"}:
            return str(value).strip()
    return ""


def _target_used_part_ids(lines, part_lookup):
    section_names = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("[PositionData") and stripped.endswith("]"):
            section_names.append(stripped[1:-1])

    part_ids = set()
    for section_name in section_names:
        try:
            rows = _read_section_rows(lines, section_name)
        except ServiceError:
            continue
        for row in rows:
            part_id = str(row.get("PARTS", "")).strip()
            if part_id in {"", "0", "-1"} or part_id not in part_lookup:
                continue
            part_name = str(part_lookup.get(part_id, {}).get("NAME", "")).strip()
            if part_name.upper() in {"", "OHM", "BLANK"}:
                continue
            part_ids.add(part_id)
    return part_ids


def _unique_target_parts(part_ids, part_lookup):
    parts = {}
    for part_id in sorted(part_ids, key=natural_sort_key):
        part_row = part_lookup.get(part_id, {})
        part_number = str(part_row.get("NAME", "")).strip()
        part_key = _part_key(part_number)
        if not part_key or part_key in parts:
            continue
        parts[part_key] = {
            "part_id": part_id,
            "part_number": part_number,
            "part_row": part_row,
        }
    return parts


def _fixed_feeder_header(lines):
    start, _ = _section_bounds(lines, "FixedFeeder")
    if start + 1 >= len(lines):
        raise ServiceError("Header [FixedFeeder] tidak ditemukan.", title="Format NPM tidak valid")
    return _split_line(lines[start + 1].strip(), "FixedFeeder", start + 2)


def _existing_fixed_assignments(fixed_rows, part_lookup, feeder_lookup):
    assignments = {}
    for row in fixed_rows:
        for assignment in _fixed_row_assignments(row, part_lookup, feeder_lookup):
            assignments.setdefault(_part_key(assignment["part_number"]), []).append(assignment)
    return assignments


def _existing_feeder_id_for_location(existing_assignments, part_key, location_code):
    for assignment in existing_assignments.get(part_key, []):
        if assignment.get("location_code") == location_code:
            return assignment.get("feeder_id", "")
    return ""


def _fixed_row_assignments(row, part_lookup, feeder_lookup):
    table, slot = _decode_pickup_unit(row.get("PU", ""))
    if not table or not slot:
        return []

    assignments = []
    for side in ("A", "B"):
        feeder_id = str(row.get(f"Feeder{side}", "")).strip()
        part_id = str(row.get(f"Parts{side}", "")).strip()
        if not _is_active(feeder_id, part_id, part_lookup):
            continue

        kind = _feeder_kind(feeder_lookup.get(feeder_id, {}))
        if kind == 2:
            position = "L" if side == "A" else "R"
            location_code = f"[{table}]{slot}{position}"
        elif kind == 4:
            location_code = f"[{table}]{slot}-{slot + 1}"
        else:
            location_code = f"[{table}]{slot}"

        assignments.append(
            {
                "part_number": part_lookup[part_id]["NAME"],
                "location_code": location_code,
                "feeder_id": feeder_id,
            }
        )
    return assignments


def _empty_fixed_row(row):
    output = dict(row)
    for side in FIXED_FEEDER_SIDES:
        output[f"Feeder{side}"] = "-1"
        output[f"Parts{side}"] = "0"
    return output


def _set_fixed_assignment(row, side, feeder_id, part_id):
    for extra_side in FIXED_FEEDER_SIDES[2:]:
        row[f"Feeder{extra_side}"] = "0"
        row[f"Parts{extra_side}"] = "0"
    row[f"Feeder{side}"] = str(feeder_id)
    row[f"Parts{side}"] = str(part_id)


def _parse_location_code(location_code):
    match = re.match(r"^\[(\d+)\](\d+)(?:-\d+)?([LR])?$", str(location_code or "").strip(), flags=re.IGNORECASE)
    if not match:
        return None
    table = int(match.group(1))
    slot = int(match.group(2))
    side = "B" if str(match.group(3) or "").upper() == "R" else "A"
    return table * 10000 + slot, side


def _default_feeder_id(part_row, feeder_lookup):
    for key in ("FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ"):
        feeder_id = str(part_row.get(key, "")).strip()
        if feeder_id not in {"", "0", "-1"} and feeder_id in feeder_lookup:
            return feeder_id
    return ""


def _replace_fixed_feeder_section(lines, fixed_header, fixed_rows):
    start, end = _section_bounds(lines, "FixedFeeder")
    newline = _detect_newline(lines)
    section_lines = [
        _strip_newline(lines[start]) + newline,
        " ".join(fixed_header) + newline,
    ]
    for row in fixed_rows:
        section_lines.append(" ".join(str(row.get(column, "")) for column in fixed_header) + newline)
    return list(lines[:start]) + section_lines + list(lines[end:])


def _section_bounds(lines, section_name):
    marker = f"[{section_name}]"
    start = None
    for index, line in enumerate(lines):
        if line.strip().lower() == marker.lower():
            start = index
            break
    if start is None:
        raise ServiceError(f"Section [{section_name}] tidak ditemukan.", title="Format NPM tidak valid")

    end = len(lines)
    for index in range(start + 1, len(lines)):
        stripped = lines[index].strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            end = index
            break
    return start, end


def _detect_newline(lines):
    for line in lines:
        if line.endswith("\r\n"):
            return "\r\n"
        if line.endswith("\n"):
            return "\n"
    return "\n"


def _strip_newline(line):
    return str(line).rstrip("\r\n")


def _part_key(value):
    return str(value or "").strip().upper()


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
    border_side = Side(style="thin", color="FFD3D3D3")
    border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    widths = {
        "A": 28,
        "B": 22,
        "C": 24,
        "D": 22,
        "E": 18,
        "F": 14,
        "G": 44,
    }
    for column_letter, width in widths.items():
        worksheet.column_dimensions[column_letter].width = width

    for cell in worksheet[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row in worksheet.iter_rows(min_row=header_row + 1):
        for cell in row:
            cell.font = default_font
            cell.border = border
            horizontal = "left" if cell.column in (1, 2, 7) else "center"
            cell.alignment = Alignment(horizontal=horizontal, vertical="center")

    worksheet.freeze_panes = f"A{header_row + 1}"
    worksheet.auto_filter.ref = f"A{header_row}:G{worksheet.max_row}"


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
