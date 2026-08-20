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
            diffs.append((circuit, "ALL", "", f"{program_row['x']}|{program_row['y']}|{program_row['angle']}|{program_row['parts']}", "ADD", "Circuit found in Program but missing from Machine file"))
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



def _load_ohm_ini_database():
    import os
    file_path = r"C:\PROGRAMMER\≪ OHM_INI ≫.xlsb"
    if not os.path.isfile(file_path):
        return {}
    
    import win32com.client
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    lib = {}
    try:
        wb = excel.Workbooks.Open(file_path, ReadOnly=True)
        ws = wb.Worksheets(2)
        data = ws.UsedRange.Value
        if data:
            for i, row in enumerate(data):
                if i == 0: continue
                if len(row) > 4:
                    part = str(row[0] or "").strip()
                    spec = str(row[1] or "").strip()
                    fd = str(row[4] or "").strip()
                    if part:
                        lib[part] = {"spec": spec, "fd": fd}
    except Exception as e:
        print(f"Failed to read OHM_INI: {e}")
    finally:
        try: wb.Close(False)
        except: pass
        excel.Quit()
    return lib

def prepare_bm221_sync(file_path, diff_results):
    try:
        with open(file_path, "r", encoding="latin-1") as handle:
            lines = handle.readlines()
    except Exception as exc:
        raise ServiceError(f"Gagal membaca file .POS:\n{exc}") from exc

    ohm_ini = None
    
    setup_lines = []
    ncdata_lines = []
    other_lines = []
    
    current_section = None
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("%"):
            current_section = stripped
            other_lines.append(line)
            continue
            
        if current_section == "%SETUP":
            setup_lines.append(line)
        elif current_section == "%NCDATA":
            ncdata_lines.append(line)
        else:
            other_lines.append(line)

    machine_circuit_to_z = {}
    
    for line in ncdata_lines:
        c_match = re.search(r"C\((.*?)\)", line)
        z_match = re.search(r"(Z\d+)", line)
        if c_match and z_match:
            machine_circuit_to_z[c_match.group(1).strip()] = z_match.group(1)
            
    part_to_z = {}
    for line in setup_lines:
        z_match = re.search(r"^(Z\d+)PC\(", line)
        pc_match = re.search(r"PC\((.*?)\)", line)
        if z_match and pc_match:
            part_to_z[pc_match.group(1).strip()] = z_match.group(1)
            
    feeder_zs = [int(z[1:]) for z in part_to_z.values() if int(z[1:]) < 100]
    tray_zs = [int(z[1:]) for z in part_to_z.values() if int(z[1:]) >= 200]
    
    max_feeder_z = max(feeder_zs) if feeder_zs else 10
    max_tray_z = max(tray_zs) if tray_zs else 200

    def get_or_create_z(part):
        nonlocal max_feeder_z, max_tray_z, ohm_ini, part_to_z, setup_lines
        if part in part_to_z:
            return part_to_z[part]
            
        if ohm_ini is None:
            ohm_ini = _load_ohm_ini_database()
            
        info = ohm_ini.get(part, {})
        fd = info.get("fd", "")
        spec_full = info.get("spec", part)
        
        is_tray = "TRAY" in fd.upper()
        
        if is_tray:
            max_tray_z += 1
            z_num = f"Z{max_tray_z}"
            spec_8 = (spec_full + "        ")[:8]
            pn_20 = (spec_full + "                    ")[:20]
            pc_15 = (part + "               ")[:15]
            new_setup = f"{z_num}PC({pc_15})PN({pn_20})ST0VO(0,0)SZ0VH0SC({spec_8})JN1\n"
        else:
            max_feeder_z = ((max_feeder_z // 5) + 1) * 5
            z_num = f"Z{max_feeder_z}"
            pc_15 = (part + "               ")[:15]
            pn_20 = (spec_full + "                    ")[:20]
            new_setup = f"{z_num}PC({pc_15})PN({pn_20})ST0VO(-300,-800)SZ0VH0SC()JN1\n"
            
        part_to_z[part] = z_num
        setup_lines.append(new_setup)
        return z_num

    replacements_made = 0
    skipped_tray_circuits = set()
    
    # Track circuits to delete
    circuits_to_delete = set()
    
    # Process diffs
    for diff in diff_results:
        circuit, field, _, program_value, diff_type, _ = diff
        
        if diff_type == "DEL":
            circuits_to_delete.add(circuit)
            replacements_made += 1
                
        elif diff_type == "CNG":
            for idx, line in enumerate(ncdata_lines):
                c_match = re.search(r"C\((.*?)\)", line)
                if c_match and c_match.group(1).strip() == circuit:
                    orig_line = ncdata_lines[idx]
                    if field == "Parts":
                        new_z = get_or_create_z(program_value)
                        ncdata_lines[idx] = re.sub(r"Z\d+", new_z, orig_line)
                        replacements_made += 1
                    elif field == "X Coordinate":
                        try:
                            val = str(int(round(float(program_value) * 1000)))
                            ncdata_lines[idx] = re.sub(r"X-?\d+", f"X{val}", orig_line)
                            replacements_made += 1
                        except: pass
                    elif field == "Y Coordinate":
                        try:
                            val = str(int(round(float(program_value) * 1000)))
                            ncdata_lines[idx] = re.sub(r"Y-?\d+", f"Y{val}", orig_line)
                            replacements_made += 1
                        except: pass
                    elif field == "Angle":
                        try:
                            val = str(int(round(float(program_value) * 100)))
                            ncdata_lines[idx] = re.sub(r"V-?\d+", f"V{val}", orig_line)
                            replacements_made += 1
                        except: pass

        elif diff_type == "ADD":
            # format: X|Y|Angle|Parts
            parts_arr = program_value.split("|")
            if len(parts_arr) == 4:
                x_val_str = str(int(round(float(parts_arr[0]) * 1000)))
                y_val_str = str(int(round(float(parts_arr[1]) * 1000)))
                v_val_str = str(int(round(float(parts_arr[2]) * 100)))
                part_val = parts_arr[3]
                
                new_z = get_or_create_z(part_val)
                # Assign MH arbitrarily or based on Tray/Feeder. 
                # For demo, let's use MH4 for Tray and MH8 for Feeder
                mh = "MH4" if int(new_z[1:]) >= 200 else "MH8"
                circuit_pad = (circuit + "        ")[:8]
                new_ncdata_line = f"N0X{x_val_str}Y{y_val_str}V{v_val_str}{new_z}{mh}C({circuit_pad})H0M000000SM0PW0BD1MN0BN0/0\n"
                ncdata_lines.append(new_ncdata_line)
                replacements_made += 1

    # Filter out deleted circuits
    final_ncdata = []
    for line in ncdata_lines:
        c_match = re.search(r"C\((.*?)\)", line)
        if c_match and c_match.group(1).strip() in circuits_to_delete:
            continue
        final_ncdata.append(line)
        
    # Reconstruct lines
    output_lines = []
    current_section = None
    for line in lines:
        if line.strip().startswith("%"):
            current_section = line.strip()
            output_lines.append(line)
            if current_section == "%SETUP":
                output_lines.extend(setup_lines)
            elif current_section == "%NCDATA":
                # Need to re-number N lines
                for idx, nc_line in enumerate(final_ncdata):
                    # Replace Nxxx with N(idx+1)
                    new_nc = re.sub(r"^N\d+", f"N{idx+1}", nc_line)
                    output_lines.append(new_nc)
        else:
            if current_section not in ["%SETUP", "%NCDATA"]:
                output_lines.append(line)

    return SyncPrepareResult(output_lines, replacements_made, skipped_tray_circuits)

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
