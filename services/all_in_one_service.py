import os
import re

from services import machine_service
from utils.sort import natural_sort_key


FILTER_NG_ONLY = "NG ONLY (Hiding Match)"
FILTER_SHOW_ALL = "SHOW ALL"
FILTER_BEDA = "BEDA DATA"
FILTER_ADD = "ADD (Hanya di TXT)"
FILTER_REMOVE = "REMOVE (Hanya di Mesin)"


FILTER_OPTIONS = [FILTER_NG_ONLY, FILTER_SHOW_ALL, FILTER_BEDA, FILTER_ADD, FILTER_REMOVE]


def classify_bulk_files(files):
    paths = {
        "npm_crb": "",
        "npm_txt": "",
        "cm602_machine": "",
        "cm602_txt": "",
        "bm_pos": "",
        "bm_txt": "",
        "bom_ori": "",
        "bom_txt": "",
    }
    scores = {key: 0 for key in paths}

    for file_path in _expand_bulk_file_candidates(files):
        name = os.path.basename(file_path).upper()
        ext = os.path.splitext(file_path)[1].lower()
        if _looks_like_cm602_machine_file(file_path):
            _set_detected_path(paths, scores, "cm602_machine", file_path, 100)
            continue

        if ext == ".crb":
            score = 95 if _looks_like_npm_crb(file_path) else 70
            _set_detected_path(paths, scores, "npm_crb", file_path, score)
        elif ext == ".pos":
            _set_detected_path(paths, scores, "bm_pos", file_path, 90)
        elif ext in [".tsv", ".csv"]:
            _set_detected_path(paths, scores, "bom_ori", file_path, 80)
        elif ext == ".txt":
            if _looks_like_cm602_machine_file(file_path):
                _set_detected_path(paths, scores, "cm602_machine", file_path, 100)
            elif _is_bom_txt_name(name) or _looks_like_bom_txt(file_path):
                _set_detected_path(paths, scores, "bom_txt", file_path, 90 if _is_bom_txt_name(name) else 55)
            elif _looks_like_program_txt(file_path):
                if _is_cm_txt_name(name):
                    _set_detected_path(paths, scores, "cm602_txt", file_path, 90)
                elif _is_bm_txt_name(name):
                    _set_detected_path(paths, scores, "bm_txt", file_path, 90)
                elif _is_npm_txt_name(name):
                    _set_detected_path(paths, scores, "npm_txt", file_path, 90)

    return paths


def _expand_bulk_file_candidates(paths):
    selected_seen = set()
    selected_files = []
    scan_dirs = []

    for raw_path in paths or []:
        if not raw_path:
            continue
        path = os.path.abspath(raw_path)
        if os.path.isdir(path):
            scan_dirs.append(path)
        elif os.path.isfile(path):
            _append_unique(selected_files, selected_seen, path)

    # Auto-import users often select only visible files from one folder, while
    # NPM exports can arrive as a folder with the .crb stored one level deeper.
    has_direct_crb = any(os.path.splitext(path)[1].lower() == ".crb" for path in selected_files)
    if selected_files and not has_direct_crb:
        parent_dirs = {os.path.dirname(path) for path in selected_files}
        scan_dirs.extend(sorted(parent_dirs, key=natural_sort_key))

    seen = set()
    candidates = []
    for path in selected_files:
        _append_unique(candidates, seen, path)

    for folder in sorted(set(scan_dirs), key=natural_sort_key):
        for root, _, filenames in os.walk(folder):
            for filename in filenames:
                candidate = os.path.join(root, filename)
                ext = os.path.splitext(candidate)[1].lower()
                if ext in {"", ".crb", ".pos", ".txt", ".tsv", ".csv"}:
                    _append_unique(candidates, seen, candidate)

    return sorted(candidates, key=lambda item: natural_sort_key(os.path.normcase(item)))


def _append_unique(items, seen, path):
    normalized = os.path.normcase(os.path.abspath(path))
    if normalized not in seen:
        seen.add(normalized)
        items.append(path)


def _set_detected_path(paths, scores, key, path, score):
    current_score = scores.get(key, 0)
    if not paths.get(key) or score > current_score:
        paths[key] = path
        scores[key] = score


def _name_tokens(name):
    stem = os.path.splitext(name.upper())[0]
    return [token for token in re.split(r"[^A-Z0-9]+", stem) if token]


def _is_cm_txt_name(name):
    tokens = _name_tokens(name)
    return "CM" in tokens or "CM602" in tokens


def _is_bm_txt_name(name):
    return "BM" in _name_tokens(name)


def _is_npm_txt_name(name):
    return "NPM" in _name_tokens(name)


def _is_bom_txt_name(name):
    tokens = _name_tokens(name)
    return "BOM" in tokens or "BOM-PG" in name.upper() or "____BOM-PG_____" in name.upper()


def parse_txt(path):
    data = {}
    try:
        with open(path, "r", errors="ignore") as handle:
            for line in handle:
                parts = line.strip().split("\t")
                if len(parts) >= 11:
                    data[parts[1].strip()] = {
                        "x": float(parts[5]),
                        "y": float(parts[6]),
                        "angle": (float(parts[7]) + 360) % 360,
                        "pn": parts[10].strip(),
                    }
    except Exception:
        pass
    return data


def parse_crb(path):
    part_map = {}
    pos_data = {}
    in_part = False
    in_pos = False
    header = []

    try:
        with open(path, "r", errors="ignore") as handle:
            lines = handle.readlines()

        for line in lines:
            line = line.strip()
            if line.startswith("[PartsData"):
                in_part = True
            elif line.startswith("["):
                in_part = False
            elif in_part and line and not line.startswith("IDNUM"):
                parts = re.split(r"\s+", line)
                if len(parts) >= 2:
                    part_map[parts[0]] = parts[1].replace('"', "")

        for line in lines:
            line = line.strip()
            if line.startswith("[PositionData"):
                in_pos = True
                header = []
            elif line.startswith("["):
                in_pos = False
            elif in_pos and line:
                if line.startswith("IDNUM"):
                    header = re.split(r"\s+", line)
                else:
                    parts = re.findall(r'"([^"]*)"|(\S+)', line)
                    parts = [value[0] if value[0] else value[1] for value in parts]
                    if header:
                        try:
                            c_i = header.index("C")
                            x_i = header.index("X")
                            y_i = header.index("Y")
                            a_i = header.index("A")
                            p_i = header.index("PARTS")
                            ref = parts[c_i].replace('"', "")
                            pos_data[ref] = {
                                "x": float(parts[x_i]),
                                "y": float(parts[y_i]),
                                "angle": (float(parts[a_i]) + 360) % 360,
                                "pn": part_map.get(parts[p_i], "UNKNOWN"),
                            }
                        except Exception:
                            pass
    except Exception:
        pass

    return pos_data


def parse_pos(path):
    part_map = {}
    pos_data = {}
    try:
        with open(path, "r", errors="ignore") as handle:
            lines = handle.readlines()

        for line in lines:
            match = re.match(r"^Z(\d+)PC\(([^)]*)\)", line.strip())
            if match:
                part_map[match.group(1)] = match.group(2).strip()

        for line in lines:
            if line.strip().startswith("N"):
                match = re.search(r"X(\d+)Y(\d+)V(\d+)Z(\d+).*?C\((.*?)\)", line)
                if match:
                    x, y, angle, z, ref = match.groups()
                    pos_data[ref.strip()] = {
                        "x": float(x) / 1000.0,
                        "y": float(y) / 1000.0,
                        "angle": (float(angle) / 100.0 + 360) % 360,
                        "pn": part_map.get(z, "UNKNOWN"),
                    }
    except Exception:
        pass
    return pos_data


def parse_bom_tsv(path):
    bom_data = {}
    try:
        with open(path, "r", errors="ignore") as handle:
            lines = handle.readlines()
            header = []
            child_idx = -1
            designator_idx = -1
            for line in lines:
                parts = line.strip("\n").split("\t")
                if not header and "Child" in parts and "Designators" in parts:
                    header = parts
                    child_idx = parts.index("Child")
                    designator_idx = parts.index("Designators")
                    continue
                if header and len(parts) > max(child_idx, designator_idx) and parts[designator_idx].strip():
                    refs = [ref.strip() for ref in parts[designator_idx].split(",") if ref.strip()]
                    for ref in refs:
                        bom_data[ref] = parts[child_idx].strip()
    except Exception:
        pass
    return bom_data


def parse_bom_txt(path):
    bom_data = {}
    try:
        with open(path, "r", errors="ignore") as handle:
            for line in handle:
                parts = line.strip().split("\t")
                if len(parts) >= 2:
                    bom_data[parts[0].strip()] = parts[1].strip()
    except Exception:
        pass
    return bom_data


def process_compare(paths):
    results = []

    if paths.get("npm_txt") and paths.get("npm_crb"):
        results.extend(get_rows("NPM", parse_txt(paths["npm_txt"]), parse_crb(paths["npm_crb"])))
    if paths.get("cm602_txt") and paths.get("cm602_machine"):
        results.extend(get_machine_service_rows("CM602", paths["cm602_machine"], paths["cm602_txt"], "CM602"))
    if paths.get("bm_txt") and paths.get("bm_pos"):
        results.extend(get_rows("BM", parse_txt(paths["bm_txt"]), parse_pos(paths["bm_pos"])))
    if paths.get("bom_txt") and paths.get("bom_ori"):
        results.extend(get_bom_rows(parse_bom_txt(paths["bom_txt"]), parse_bom_tsv(paths["bom_ori"])))

    return results


def get_rows(system_name, txt_data, machine_data):
    all_refs = sorted(list(set(txt_data.keys()) | set(machine_data.keys())))
    rows = []
    for ref in all_refs:
        if str(ref).strip().lower() in ["1", "2", "ohm", ""]:
            continue

        in_txt = ref in txt_data
        in_machine = ref in machine_data
        if in_txt and in_machine:
            target = txt_data[ref]
            source = machine_data[ref]
            source_str = f"{source['pn']} | {source['x']:.3f}/{source['y']:.3f} | {source['angle']:.3f}"
            target_str = f"{target['pn']} | {target['x']:.3f}/{target['y']:.3f} | {target['angle']:.3f}"
            diffs = []
            if target["pn"].upper() != source["pn"].upper():
                diffs.append("P/N")
            if abs(target["x"] - source["x"]) > 0.05 or abs(target["y"] - source["y"]) > 0.05:
                diffs.append("KOORDINAT")
            if target["angle"] != source["angle"]:
                diffs.append("ANGLE")

            is_match = len(diffs) == 0
            status = "MATCH" if is_match else "BEDA " + ", ".join(diffs)
            rows.append(_record(ref, system_name, status, source_str, target_str, "match" if is_match else "error"))
        elif in_txt:
            target = txt_data[ref]
            target_str = f"{target['pn']} | {target['x']:.3f}/{target['y']:.3f} | {target['angle']:.3f}"
            rows.append(_record(ref, system_name, "ADD (Hanya di TXT)", "-", target_str, "add"))
        else:
            source = machine_data[ref]
            source_str = f"{source['pn']} | {source['x']:.3f}/{source['y']:.3f} | {source['angle']:.3f}"
            rows.append(_record(ref, system_name, "REMOVE (Hanya di Mesin)", source_str, "-", "remove"))
    return rows


def get_machine_service_rows(system_name, machine_path, program_path, machine_type):
    machine_df = machine_service.load_machine_file(machine_path, machine_type)
    program_df = machine_service.load_program_file(program_path, machine_type)
    diff_results = machine_service.compare_machine(machine_df, program_df)

    machine_dict = {row["circuit"]: row for _, row in machine_df.iterrows()}
    program_dict = {row["circuit"]: row for _, row in program_df.iterrows()}
    all_refs = sorted(set(machine_dict.keys()) | set(program_dict.keys()), key=natural_sort_key)
    diff_type_by_ref = _machine_diff_type_by_ref(diff_results)
    diff_reasons_by_ref = _machine_diff_reasons_by_ref(diff_results)

    rows = []
    for ref in all_refs:
        if str(ref).strip().lower() in ["1", "2", "ohm", "ohm1", ""]:
            continue

        machine_row = machine_dict.get(ref)
        program_row = program_dict.get(ref)
        if machine_row is not None and program_row is not None:
            diff_type = diff_type_by_ref.get(ref)
            if diff_type is None:
                status = "MATCH"
            else:
                reasons = diff_reasons_by_ref.get(ref, [])
                status = "BEDA " + ", ".join(reasons) if reasons else "BEDA DATA"
            rows.append(
                _record(
                    ref,
                    system_name,
                    status,
                    _machine_row_text(machine_row),
                    _machine_row_text(program_row),
                    "match" if status == "MATCH" else "error",
                )
            )
        elif program_row is not None:
            rows.append(_record(ref, system_name, "ADD (Hanya di TXT)", "-", _machine_row_text(program_row), "add"))
        else:
            rows.append(_record(ref, system_name, "REMOVE (Hanya di Mesin)", _machine_row_text(machine_row), "-", "remove"))

    return rows


def get_bom_rows(txt_data, ori_data):
    all_refs = sorted(list(set(txt_data.keys()) | set(ori_data.keys())))
    rows = []
    for ref in all_refs:
        if str(ref).strip().lower() in ["1", "2", "ohm", ""]:
            continue

        in_txt = ref in txt_data
        in_ori = ref in ori_data
        if in_txt and in_ori:
            target_pn = txt_data[ref]
            source_pn = ori_data[ref]
            is_match = target_pn.upper() == source_pn.upper()
            status = "MATCH" if is_match else "BEDA P/N"
            rows.append(_record(ref, "BOM", status, source_pn, target_pn, "match" if is_match else "error"))
        elif in_txt:
            rows.append(_record(ref, "BOM", "ADD (Hanya di TXT)", "-", txt_data[ref], "add"))
        else:
            rows.append(_record(ref, "BOM", "REMOVE (Hanya di Ori)", ori_data[ref], "-", "remove"))
    return rows


def _machine_diff_type_by_ref(diff_results):
    output = {}
    for diff in diff_results:
        ref = diff[0]
        diff_type = diff[4]
        current = output.get(ref)
        if current == "CNG":
            continue
        if diff_type == "CNG" or current is None:
            output[ref] = diff_type
    return output


def _machine_diff_reasons_by_ref(diff_results):
    output = {}
    for diff in diff_results:
        ref = diff[0]
        field = diff[1]
        diff_type = diff[4]
        if diff_type == "CNG":
            if ref not in output:
                output[ref] = []
            
            if field in ["X Coordinate", "Y Coordinate"]:
                if "KOORDINAT" not in output[ref]:
                    output[ref].append("KOORDINAT")
            elif field == "Angle":
                if "ANGLE" not in output[ref]:
                    output[ref].append("ANGLE")
            elif field == "Parts":
                if "P/N" not in output[ref]:
                    output[ref].append("P/N")
    return output


def _machine_row_text(row):
    if row is None:
        return "-"
    return (
        f"{row.get('parts', '')} | "
        f"{_format_float(row.get('x', ''))}/{_format_float(row.get('y', ''))} | "
        f"{_format_float(row.get('angle', ''))}"
    )


def _format_float(value):
    try:
        return f"{float(value):.3f}"
    except (TypeError, ValueError):
        return str(value)


def _looks_like_cm602_machine_file(path):
    try:
        with open(path, "r", encoding="latin-1", errors="ignore") as handle:
            sample = handle.read(120000)
    except Exception:
        return False
    return "[BlockData]" in sample and "[PartsData]" in sample


def _looks_like_npm_crb(path):
    try:
        with open(path, "r", encoding="latin-1", errors="ignore") as handle:
            sample = handle.read(120000)
    except Exception:
        return False
    return "[PartsData" in sample and "[PositionData" in sample


def _looks_like_program_txt(path):
    try:
        with open(path, "r", encoding="latin-1", errors="ignore") as handle:
            for line in handle:
                parts = line.rstrip("\n").split("\t")
                if len(parts) >= 11:
                    return True
    except Exception:
        return False
    return False


def _looks_like_bom_txt(path):
    bom_rows = 0
    program_rows = 0
    try:
        with open(path, "r", encoding="latin-1", errors="ignore") as handle:
            for line in handle:
                parts = [part.strip() for part in line.rstrip("\n").split("\t")]
                if not any(parts):
                    continue
                if len(parts) >= 11:
                    program_rows += 1
                elif len(parts) >= 2 and _looks_like_ref(parts[0]) and parts[1]:
                    bom_rows += 1
                if bom_rows >= 3 or program_rows >= 3:
                    break
    except Exception:
        return False
    return bom_rows >= 3 and program_rows == 0


def _looks_like_ref(value):
    return bool(re.match(r"^[A-Z]{1,4}\d", str(value).strip().upper()))


def filter_results(results, filter_type):
    filtered = []
    for result in results:
        status = result["status"]
        show = False
        if filter_type == FILTER_NG_ONLY:
            show = status != "MATCH"
        elif filter_type == FILTER_SHOW_ALL:
            show = True
        elif filter_type == FILTER_BEDA and "BEDA" in status:
            show = True
        elif filter_type == FILTER_ADD and "ADD" in status:
            show = True
        elif filter_type == FILTER_REMOVE and "REMOVE" in status:
            show = True

        if show:
            filtered.append(result)
    return filtered


def _record(ref, system, status, source, target, tag):
    diff_keys = []
    if status != "MATCH":
        diff_keys = ["status", "source", "target"]
    return {
        "ref": ref,
        "system": system,
        "status": status,
        "source": source,
        "target": target,
        "tag": tag,
        "_diff_keys": diff_keys,
    }

