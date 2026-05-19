import os
import re
import shlex
from dataclasses import dataclass
from datetime import datetime

import pandas as pd

from services.errors import ServiceError
from utils.encoding import ENCODINGS
from utils.sort import natural_sort_key


@dataclass
class SyncPrepareResult:
    content: list
    replacements_made: int
    skipped_tray_circuits: set


IGNORED_CIRCUITS = {"ohm", "ohm1"}


def _is_ignored_circuit(circuit):
    return not circuit or circuit.strip().lower() in IGNORED_CIRCUITS


def load_machine_file(file_path, machine_type):
    try:
        with open(file_path, "r", encoding="latin-1") as handle:
            content = handle.readlines()
    except Exception as exc:
        raise ServiceError(f"Gagal membaca file CRB:\n{exc}") from exc

    position_rows = []

    if machine_type == "NPM":
        parts_map = {}
        current_section = None
        for line in content:
            line = line.strip()
            if not line:
                continue
            if line.startswith("["):
                current_section = line
                continue

            if current_section == "[PartsData]":
                cols = line.split()
                if len(cols) > 1 and cols[0].isdigit():
                    parts_map[cols[0]] = cols[1].strip('"')
            elif current_section == "[PositionData<1>]":
                cols = line.split()
                if len(cols) > 24 and cols[0].isdigit():
                    circuit = cols[24].strip('"').strip()
                    if _is_ignored_circuit(circuit):
                        continue
                    position_rows.append(
                        {
                            "circuit": circuit,
                            "x": cols[2],
                            "y": cols[3],
                            "angle": cols[4],
                            "partno": cols[5],
                        }
                    )
        for row in position_rows:
            row["parts"] = parts_map.get(row["partno"], "Unknown")

    elif machine_type == "BM221":
        parts_map = {}
        current_section = None
        for line in content:
            line = line.strip()
            if not line:
                continue
            if line.startswith("%"):
                current_section = line
                continue

            if current_section == "%SETUP":
                match = re.search(r"(Z\d+)PC\((.*?)\)", line)
                if match:
                    parts_map[match.group(1)] = match.group(2).strip()
            elif current_section == "%NCDATA":
                x_match = re.search(r"X(\d+)", line)
                y_match = re.search(r"Y(\d+)", line)
                v_match = re.search(r"V(\d+)", line)
                z_match = re.search(r"(Z\d+)", line)
                c_match = re.search(r"C\((.*?)\)", line)

                if x_match and y_match and z_match and c_match:
                    x_val = float(x_match.group(1)) / 1000.0
                    y_val = float(y_match.group(1)) / 1000.0
                    angle_val = 0.0
                    if v_match:
                        angle_val = float(v_match.group(1)) / 100.0

                    position_rows.append(
                        {
                            "circuit": c_match.group(1).strip(),
                            "x": f"{x_val:.3f}",
                            "y": f"{y_val:.3f}",
                            "angle": f"{angle_val:.1f}",
                            "partno": z_match.group(1),
                        }
                    )
        for row in position_rows:
            row["parts"] = parts_map.get(row["partno"], "Unknown")

    elif machine_type == "CM602":
        parts_map = {}
        current_section = None
        for line in content:
            line = line.strip()
            if not line:
                continue
            if line.startswith("["):
                current_section = line
                continue

            if current_section == "[PartsData]":
                try:
                    cols = shlex.split(line)
                    if len(cols) > 1 and cols[0].isdigit():
                        parts_map[cols[0]] = cols[1].strip()
                except ValueError:
                    pass
            elif current_section == "[BlockData]":
                try:
                    cols = shlex.split(line)
                    if len(cols) >= 8 and cols[0].isdigit():
                        position_rows.append(
                            {
                                "circuit": cols[7].strip(),
                                "x": cols[2],
                                "y": cols[3],
                                "angle": cols[4],
                                "partno": cols[5],
                            }
                        )
                except ValueError:
                    pass
        for row in position_rows:
            row["parts"] = parts_map.get(row["partno"], "Unknown")

    df = pd.DataFrame(position_rows)
    if df.empty:
        raise ServiceError("No valid position data found in machine file.", title="Warning")

    return df


def load_program_file(file_path, machine_type):
    df = None
    last_error = None
    for encoding in ENCODINGS:
        try:
            df = pd.read_csv(file_path, sep="\t", header=None, dtype=str, keep_default_na=False, encoding=encoding)
            break
        except UnicodeDecodeError as exc:
            last_error = exc

    if df is None:
        raise ServiceError(f"Gagal membaca file TXT:\n{last_error}")

    if df.shape[1] < 11:
        raise ServiceError("Format file TXT tidak sesuai (kurang kolom).")

    extracted = []
    x_idx = 3 if machine_type == "CM602" else 5
    y_idx = 4 if machine_type == "CM602" else 6

    for _, row in df.iterrows():
        circuit = str(row[1]).strip()
        if not circuit:
            continue
        extracted.append(
            {
                "circuit": circuit,
                "partno": str(row[0]).strip(),
                "x": str(row[x_idx]).strip(),
                "y": str(row[y_idx]).strip(),
                "angle": str(row[7]).strip(),
                "parts": str(row[10]).strip(),
            }
        )

    result_df = pd.DataFrame(extracted)
    if result_df.empty:
        raise ServiceError("No valid program data found.", title="Warning")

    return result_df


def compare_machine(machine_df, program_df):
    machine_dict = {row["circuit"]: row for _, row in machine_df.iterrows()}
    program_dict = {row["circuit"]: row for _, row in program_df.iterrows()}
    all_circuits = sorted(set(machine_dict.keys()) | set(program_dict.keys()), key=natural_sort_key)
    diffs = []

    def normalize_num(value):
        try:
            return float(value)
        except Exception:
            return None

    def normalize_angle(value):
        try:
            return float(value) % 360
        except Exception:
            return None

    fields_to_check = [
        ("x", "X Coordinate", normalize_num, 0.001),
        ("y", "Y Coordinate", normalize_num, 0.001),
        ("angle", "Angle", normalize_angle, 0.01),
        ("parts", "Parts", lambda value: str(value).strip(), None),
    ]

    for circuit in all_circuits:
        machine_row = machine_dict.get(circuit)
        program_row = program_dict.get(circuit)

        if machine_row is not None and program_row is not None:
            for key, label, norm_func, tolerance in fields_to_check:
                machine_value = norm_func(machine_row[key])
                program_value = norm_func(program_row[key])

                if tolerance is not None:
                    if machine_value is not None and program_value is not None:
                        is_match = abs(machine_value - program_value) <= tolerance
                    else:
                        is_match = machine_value is None and program_value is None
                else:
                    is_match = machine_value == program_value

                if not is_match:
                    diffs.append(
                        (
                            circuit,
                            label,
                            machine_row[key],
                            program_row[key],
                            "CNG",
                            f"Mismatch in {label}",
                        )
                    )
        elif machine_row is None:
            diffs.append((circuit, "ALL", "", "Present", "ADD", "Circuit found in Program but missing from Machine file"))
        else:
            diffs.append((circuit, "ALL", "Present", "", "DEL", "Circuit found in Machine but missing from Program file"))

    return diffs


def build_machine_diff_preview(machine_df, program_df, diff_results):
    machine_dict = {row["circuit"]: row for _, row in machine_df.iterrows()}
    program_dict = {row["circuit"]: row for _, row in program_df.iterrows()}
    type_by_circuit = {}
    diff_keys_by_circuit = {}

    for diff in diff_results:
        circuit = diff[0]
        field = diff[1]
        diff_type = diff[4]
        current = type_by_circuit.get(circuit)
        if current == "CNG":
            pass
        elif diff_type == "CNG" or current is None:
            type_by_circuit[circuit] = diff_type
        diff_keys_by_circuit.setdefault(circuit, set()).update(_diff_keys_for_field(field, diff_type))

    ordered_circuits = sorted(type_by_circuit.keys(), key=natural_sort_key)
    machine_rows = []
    program_rows = []

    for circuit in ordered_circuits:
        diff_type = type_by_circuit[circuit]
        machine_row = machine_dict.get(circuit)
        program_row = program_dict.get(circuit)

        machine_type_label = "Del" if diff_type == "DEL" else ""
        program_type_label = "Mod" if diff_type == "CNG" else ("Add" if diff_type == "ADD" else "")
        diff_keys = sorted(diff_keys_by_circuit.get(circuit, []))

        machine_rows.append(_preview_record(circuit, machine_row, machine_type_label, diff_keys))
        program_rows.append(_preview_record(circuit, program_row, program_type_label, diff_keys))

    return machine_rows, program_rows


def _diff_keys_for_field(field, diff_type):
    if diff_type in ["ADD", "DEL"]:
        return {"circuit", "x", "y", "angle", "partno", "parts"}
    return {
        "X Coordinate": {"x"},
        "Y Coordinate": {"y"},
        "Angle": {"angle"},
        "Parts": {"parts"},
    }.get(field, set())


def _preview_record(circuit, row, type_label, diff_keys):
    if row is None:
        return {
            "circuit": circuit,
            "x": "",
            "y": "",
            "angle": "",
            "partno": "",
            "parts": "",
            "type": type_label,
            "_diff_keys": diff_keys,
        }

    return {
        "circuit": row.get("circuit", circuit),
        "x": row.get("x", ""),
        "y": row.get("y", ""),
        "angle": row.get("angle", ""),
        "partno": row.get("partno", ""),
        "parts": row.get("parts", ""),
        "type": type_label,
        "_diff_keys": diff_keys,
    }


def export_machine_results(diff_results, file_path):
    pd.DataFrame(
        diff_results,
        columns=["Circuit No", "Field", "Machine Value", "Program Value", "Type", "Description"],
    ).to_excel(file_path, index=False)


def export_machine_preview(machine_rows, program_rows, file_path):
    rows = []
    for machine_row, program_row in zip(machine_rows, program_rows):
        rows.append(
            [
                machine_row.get("circuit", ""),
                machine_row.get("x", ""),
                machine_row.get("y", ""),
                machine_row.get("angle", ""),
                machine_row.get("partno", ""),
                machine_row.get("parts", ""),
                machine_row.get("type", ""),
                program_row.get("circuit", ""),
                program_row.get("x", ""),
                program_row.get("y", ""),
                program_row.get("angle", ""),
                program_row.get("partno", ""),
                program_row.get("parts", ""),
                program_row.get("type", ""),
            ]
        )

    pd.DataFrame(
        rows,
        columns=[
            "Circuit No",
            "X Coordinate",
            "Y Coordinate",
            "Angle",
            "Parts Number",
            "Parts",
            "Type",
            "Circuit No",
            "X Coordinate",
            "Y Coordinate",
            "Angle",
            "Parts Number",
            "Parts",
            "Type",
        ],
    ).to_excel(file_path, index=False)


def prepare_bm221_sync(file_path, diff_results):
    try:
        with open(file_path, "r", encoding="latin-1") as handle:
            content = handle.readlines()
    except Exception as exc:
        raise ServiceError(f"Gagal membaca file .POS:\n{exc}") from exc

    machine_circuit_to_z = {}
    setup_z_to_line_idx = {}
    ncdata_circuit_to_line_idx = {}

    current_section = None
    for i, line in enumerate(content):
        stripped = line.strip()
        if stripped.startswith("%"):
            current_section = stripped
            continue

        if current_section == "%SETUP":
            match = re.search(r"^(Z\d+)PC\(", stripped)
            if match:
                setup_z_to_line_idx[match.group(1)] = i
        elif current_section == "%NCDATA":
            c_match = re.search(r"C\((.*?)\)", stripped)
            z_match = re.search(r"(Z\d+)", stripped)
            if c_match and z_match:
                circuit = c_match.group(1).strip()
                machine_circuit_to_z[circuit] = z_match.group(1)
                ncdata_circuit_to_line_idx[circuit] = i

    skipped_tray_circuits = set()
    replacements_made = 0

    for diff in diff_results:
        circuit = diff[0]
        field = diff[1]
        program_value = diff[3]
        diff_type = diff[4]

        if diff_type != "CNG":
            continue

        if field == "Parts":
            machine_z = machine_circuit_to_z.get(circuit)
            if not machine_z:
                continue

            machine_z_num = int(machine_z[1:]) if machine_z[1:].isdigit() else 0
            if machine_z_num >= 200:
                skipped_tray_circuits.add(circuit)
                continue

            line_idx = setup_z_to_line_idx.get(machine_z)
            if line_idx is not None:
                original_line = content[line_idx]
                padded = str(program_value).ljust(15)
                content[line_idx] = re.sub(r"PC\([^)]*\)", f"PC({padded})", original_line)
                replacements_made += 1
        elif field in ["X Coordinate", "Y Coordinate", "Angle"]:
            line_idx = ncdata_circuit_to_line_idx.get(circuit)
            if line_idx is None:
                continue

            original_line = content[line_idx]
            try:
                if field == "X Coordinate":
                    value = str(int(round(float(program_value) * 1000)))
                    content[line_idx] = re.sub(r"X-?\d+", f"X{value}", original_line)
                elif field == "Y Coordinate":
                    value = str(int(round(float(program_value) * 1000)))
                    content[line_idx] = re.sub(r"Y-?\d+", f"Y{value}", original_line)
                elif field == "Angle":
                    value = str(int(round(float(program_value) * 100)))
                    content[line_idx] = re.sub(r"V-?\d+", f"V{value}", original_line)
                replacements_made += 1
            except ValueError:
                pass

    return SyncPrepareResult(content, replacements_made, skipped_tray_circuits)


def write_pos_file(content, save_path):
    try:
        with open(save_path, "w", encoding="latin-1") as handle:
            handle.writelines(content)
    except Exception as exc:
        raise ServiceError(f"Gagal menyimpan file:\n{exc}") from exc


def machine_history_entry(machine_file, program_file, diff_results):
    return {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "txt_file": f"Machine: {machine_file}",
        "tsv_file": f"Program: {program_file}",
        "add_count": sum(1 for item in diff_results if item[4] == "ADD"),
        "cng_count": sum(1 for item in diff_results if item[4] == "CNG"),
        "del_count": sum(1 for item in diff_results if item[4] == "DEL"),
        "results": diff_results,
    }
