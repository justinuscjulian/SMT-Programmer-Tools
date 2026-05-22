import io
import os
from dataclasses import dataclass
from datetime import datetime

import pandas as pd

from services.errors import DuplicateCircuitError, ServiceError
from utils.encoding import ENCODINGS, read_lines_with_fallback
from utils.sort import natural_sort_key


@dataclass
class RawBomResult:
    dataframe: pd.DataFrame
    chassis_pn: str
    pcb_pn: str
    timestamp: str


def _check_duplicate_circuits(df, column):
    duplicate_circuits = df[df.duplicated(column, keep=False)][column].unique()
    if len(duplicate_circuits) > 0:
        raise DuplicateCircuitError(duplicate_circuits)


def _sort_by_circuit(df, column="Circuit"):
    ordered_index = sorted(range(len(df)), key=lambda idx: natural_sort_key(df.iloc[idx][column]))
    return df.iloc[ordered_index].reset_index(drop=True)


def load_reference_txt(file_path, check_duplicate_circuits=False):
    df = None
    for encoding in ENCODINGS:
        try:
            df = pd.read_csv(
                file_path,
                sep="\t",
                header=None,
                dtype=str,
                keep_default_na=False,
                encoding=encoding,
            )
            break
        except UnicodeDecodeError:
            continue

    if df is None:
        raise ServiceError("Encoding error.")
    if df.shape[1] < 2:
        raise ServiceError("Format file TXT tidak sesuai (kurang kolom).")

    df = df.iloc[:, :2]
    df.columns = ["Circuit", "PartNo"]
    df["Circuit"] = df["Circuit"].str.strip()
    df["PartNo"] = df["PartNo"].str.strip()
    df = df[df["Circuit"] != ""].reset_index(drop=True)
    if check_duplicate_circuits:
        _check_duplicate_circuits(df, "Circuit")
    return df


def load_raw_bom(file_path, check_duplicate_circuits=False):
    ext = os.path.splitext(file_path)[1].lower()
    raw_df = None

    if ext in [".xlsx", ".xls"]:
        try:
            raw_df = pd.read_excel(file_path, header=None, dtype=str)
        except Exception as exc:
            raise ServiceError(f"Excel error:\n{exc}") from exc
    else:
        try:
            file_content, used_enc = read_lines_with_fallback(file_path)
        except Exception as exc:
            raise ServiceError("Encoding error.") from exc

        header_idx = -1
        header_cols_count = 0
        for i, line in enumerate(file_content):
            cols = line.strip("\n").split("\t")
            cols_lower = [c.strip().lower() for c in cols]
            if "parent" in cols_lower and "child" in cols_lower and "designators" in cols_lower:
                header_idx = i
                header_cols_count = len(cols)
                break
            if cols and cols[0].strip().lower() == "head" and len(cols) > 5:
                header_idx = i
                header_cols_count = len(cols)
                break

        if header_idx != -1:
            valid_lines = [file_content[header_idx]]
            for line in file_content[header_idx + 1 :]:
                cols = line.strip("\n").split("\t")
                if len(cols) == header_cols_count:
                    valid_lines.append(line)
                elif len(cols) > header_cols_count:
                    merged_cols = cols[: header_cols_count - 1]
                    merged_cols.append(" ".join(cols[header_cols_count - 1 :]))
                    valid_lines.append("\t".join(merged_cols) + "\n")
            raw_df = pd.read_csv(
                io.StringIO("".join(valid_lines)),
                sep="\t",
                header=None,
                dtype=str,
                keep_default_na=False,
            )
        else:
            raw_df = pd.read_csv(
                file_path,
                sep="\t",
                header=None,
                dtype=str,
                encoding=used_enc,
                on_bad_lines="skip",
            )

    if raw_df is None:
        raise ServiceError("Raw BOM tidak berhasil dibaca.")

    header_row_idx = 0
    for idx, row in raw_df.iterrows():
        row_vals = [str(x).strip().lower() for x in row]
        if "parent" in row_vals and "child" in row_vals and "designators" in row_vals:
            header_row_idx = idx
            break
        if str(row.iloc[0]).strip().lower() == "head" and len(row_vals) > 5:
            header_row_idx = idx
            break

    headers = raw_df.iloc[header_row_idx].fillna("Unnamed").astype(str).str.strip()
    data = raw_df.iloc[header_row_idx + 1 :].reset_index(drop=True)
    data.columns = headers
    df = data.copy()
    df.fillna("", inplace=True)

    col_map = {str(col).lower().strip(): col for col in df.columns}
    col_parent = col_map.get("parent")
    col_parent_desc = col_map.get("parent description")
    col_level = col_map.get("level")

    chassis_pn = _extract_root_parent_part_number(df, col_parent, col_level)
    if chassis_pn == "Unknown" and col_parent and col_parent_desc:
        fallback_pn = ""
        for _, row in df.iterrows():
            pdesc = str(row[col_parent_desc]).lower().replace(" ", "")
            ppart = str(row[col_parent]).strip()
            if "orptotalassembly" in pdesc:
                fallback_pn = fallback_pn or ppart
                continue
            if "bprtotalassembly" in pdesc:
                fallback_pn = fallback_pn or ppart
                continue
            if "chassisassembly" in pdesc:
                fallback_pn = fallback_pn or ppart
                continue
            if "pcbassembly,main" in pdesc:
                fallback_pn = fallback_pn or ppart
                continue

        chassis_pn = fallback_pn or "Unknown"

    col_child = col_map.get("child")
    col_designators = col_map.get("designators")
    col_spec = col_map.get("specification")
    col_desc = col_map.get("description")
    if not col_child or not col_designators:
        raise ServiceError("Header Child/Designators not found!")

    pcb_pn = ""
    if col_desc:
        for _, row in df.iterrows():
            if str(row[col_desc]).lower().replace(" ", "") in ["pcb", "pcb,main"]:
                pcb_pn = str(row[col_child]).strip()
                break

    output_rows = []
    for _, row in df.iterrows():
        raw_desig = str(row[col_designators]).strip()
        if not raw_desig or raw_desig.lower() == "nan":
            continue

        designators = [d.strip() for d in raw_desig.split(",") if d.strip()]
        side_val = ""
        if col_parent_desc:
            pdesc = str(row[col_parent_desc]).lower()
            if "orp insert pcb assembly" in pdesc:
                side_val = "ORP"
            elif "upper smt pcb assembly" in pdesc or "s/w package" in pdesc:
                side_val = "TOP"
            elif "bpr insert pcb assembly" in pdesc:
                side_val = "BPR"

        spec_val = ""
        if col_spec:
            raw_spec = str(row[col_spec]).strip()
            if raw_spec and raw_spec.lower() != "nan":
                spec_val = raw_spec[:60]

        child_val = str(row[col_child]).strip()
        for designator in designators:
            output_rows.append(
                {
                    "Chassis": chassis_pn,
                    "Circuit": designator,
                    "PartNo": child_val,
                    "Spec": spec_val,
                    "Side": side_val,
                }
            )

    if not output_rows:
        raise ServiceError("Tidak ada designator valid yang ditemukan.")

    result_df = pd.DataFrame(output_rows)
    result_df = _sort_by_circuit(result_df, "Circuit")
    if check_duplicate_circuits:
        _check_duplicate_circuits(result_df, "Circuit")

    return RawBomResult(
        dataframe=result_df,
        chassis_pn=chassis_pn,
        pcb_pn=pcb_pn,
        timestamp=datetime.now().strftime("%d/%m/%Y %H:%M"),
    )


def _extract_root_parent_part_number(df, col_parent, col_level):
    if not col_parent:
        return "Unknown"

    if col_level:
        for _, row in df.iterrows():
            level = str(row[col_level]).strip()
            parent = str(row[col_parent]).strip()
            if parent and level in {"1", "1.0"}:
                return parent

    for _, row in df.iterrows():
        parent = str(row[col_parent]).strip()
        if parent:
            return parent

    return "Unknown"


def compare_bom(reference_df, raw_df):
    txt_dict = {}
    tsv_dict = {}
    side_dict = {}

    for _, row in reference_df.iterrows():
        circuit = row["Circuit"]
        if circuit:
            txt_dict.setdefault(circuit, set()).add(row["PartNo"])

    for _, row in raw_df.iterrows():
        circuit = row["Circuit"]
        if circuit:
            tsv_dict.setdefault(circuit, set()).add(row["PartNo"])
            side_dict[circuit] = row["Side"]

    all_circuits = sorted(set(txt_dict.keys()) | set(tsv_dict.keys()), key=natural_sort_key)
    diffs = []

    for circuit in all_circuits:
        txt_parts = txt_dict.get(circuit, set())
        tsv_parts = tsv_dict.get(circuit, set())
        side_val = side_dict.get(circuit, "")

        if txt_parts == tsv_parts:
            continue

        txt_str = ", ".join(sorted(txt_parts)) if txt_parts else ""
        tsv_str = ", ".join(sorted(tsv_parts)) if tsv_parts else ""

        if txt_parts and tsv_parts:
            diff_type = "CNG"
            desc = "Part number mismatch on identical RefDes identifier."
        elif not txt_parts:
            diff_type = "ADD"
            desc = "Component found in Raw file but missing from Reference."
        else:
            diff_type = "DEL"
            desc = "RefDes present in Master Reference but deleted in latest Raw import."

        diffs.append((circuit, side_val, txt_str, tsv_str, diff_type, desc))

    type_order = {"ADD": 1, "CNG": 2, "DEL": 3}
    diffs.sort(key=lambda item: (type_order.get(item[4], 99), natural_sort_key(item[0])))
    return diffs


def export_bom_results(diff_results, file_path):
    pd.DataFrame(
        diff_results,
        columns=["Circuit No", "Side", "Part (Reference)", "Part (Source)", "Type", "Description"],
    ).to_excel(file_path, index=False)
