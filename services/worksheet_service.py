from dataclasses import dataclass
import os

import pandas as pd

from services.errors import ServiceError


@dataclass
class WorksheetResult:
    ng_id: list
    ng_qty: list
    matches: list


def run_worksheet_compare(file_excel, file_crb):
    file_excel = file_excel.strip().replace('"', "").replace("'", "")
    file_crb = file_crb.strip().replace('"', "").replace("'", "")

    if not os.path.exists(file_excel):
        raise ServiceError(f"File Excel ga ketemu gan!\n{file_excel}")
    if not os.path.exists(file_crb):
        raise ServiceError(f"File Mesin (CRB) ga ketemu gan!\n{file_crb}")

    excel_dict = {}

    try:
        if file_excel.lower().endswith(".csv"):
            df_raw = pd.read_csv(file_excel, header=None)
        else:
            ext = os.path.splitext(file_excel)[1].lower()
            engine = "xlrd" if ext == ".xls" else "openpyxl"
            xls = pd.ExcelFile(file_excel, engine=engine)
            try:
                df_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)
            finally:
                xls.close()

        part_col = 1
        qty_col = 2
        id_col = 6
        start_row = 3

        for row_index in range(start_row, len(df_raw)):
            part_val = str(df_raw.iloc[row_index, part_col]).strip()
            id_val = str(df_raw.iloc[row_index, id_col]).strip()

            if id_val.endswith(".0"):
                id_val = id_val[:-2]
            if part_val.endswith(".0"):
                part_val = part_val[:-2]

            if id_val.upper() in ["NAN", "NONE", ""] or part_val.upper() in ["NAN", "NONE", ""]:
                continue
            if id_val in ["1401", "1402"]:
                continue

            qty_val = 1
            raw_qty = str(df_raw.iloc[row_index, qty_col]).strip()
            if raw_qty.upper() not in ["NAN", "NONE", ""]:
                try:
                    qty_val = int(float(raw_qty))
                except ValueError:
                    qty_val = 1

            excel_dict[id_val] = {"PART": part_val, "QTY": qty_val}
    except Exception as exc:
        raise ServiceError(f"Gagal membaca Excel: {exc}") from exc

    parts_mapping = {}
    crb_dict = {}

    try:
        with open(file_crb, "r", encoding="utf-8", errors="ignore") as file:
            lines = file.readlines()

        in_parts_data = False
        for line in lines:
            line = line.strip()
            if line.startswith("[PartsData]"):
                in_parts_data = True
                continue
            if line.startswith("[") and in_parts_data:
                in_parts_data = False

            if in_parts_data and line and not line.startswith("IDNUM"):
                cols = line.split()
                if len(cols) > 2:
                    parts_idnum = cols[0]
                    part_name = cols[1].replace('"', "")
                    parts_mapping[parts_idnum] = part_name

        in_pos_data = False
        for line in lines:
            line = line.strip()
            if line.startswith("[PositionData"):
                in_pos_data = True
                continue
            if line.startswith("[") and in_pos_data:
                in_pos_data = False

            if in_pos_data and line and not line.startswith("IDNUM"):
                cols = line.split()
                if len(cols) > 15:
                    placement_id = cols[0]
                    parts_id = cols[5]
                    pu_val = cols[14]
                    side_val = cols[15]
                    actual_part_name = parts_mapping.get(parts_id, "")

                    if placement_id in ["1", "2"] or actual_part_name.upper() in ["OHM", "BLANK", ""]:
                        continue

                    try:
                        pu_int = int(pu_val)
                        side_int = int(side_val)

                        if pu_int >= 100000:
                            block = pu_int // 1000
                            slot = pu_int % 1000
                            if block == 101:
                                table = 10
                            elif block == 102:
                                table = 11
                            else:
                                table = block - 91
                        else:
                            table = pu_int // 10000
                            slot = pu_int % 10000

                        if 1 <= table <= 20:
                            if side_int == 0:
                                feed_id = str(table * 100 + (slot * 2) - 1)
                            else:
                                feed_id = str(table * 100 + (slot * 2) - (2 - side_int))
                        else:
                            feed_id = str(pu_val)
                    except ValueError:
                        feed_id = pu_val

                    if feed_id in ["1401", "1402"]:
                        continue

                    if feed_id not in crb_dict:
                        crb_dict[feed_id] = {"PART": actual_part_name, "QTY": 0}
                    crb_dict[feed_id]["QTY"] += 1
    except Exception as exc:
        raise ServiceError(f"Gagal membaca CRB: {exc}") from exc

    all_ids = set(excel_dict.keys()).union(set(crb_dict.keys()))
    ng_id = []
    ng_qty = []
    matches = []

    for feed_id in sorted(all_ids, key=lambda value: int(value) if str(value).isdigit() else value):
        if feed_id in ["1401", "1402"]:
            continue

        ex_data = excel_dict.get(feed_id)
        crb_data = crb_dict.get(feed_id)

        is_table_10_or_11 = False
        if str(feed_id).isdigit():
            table_num = int(feed_id) // 100
            if table_num in [10, 11]:
                is_table_10_or_11 = True

        if not ex_data:
            ng_id.append(f"[NG] ID {feed_id} ada di Mesin ('{crb_data['PART']}'), tapi GAADA di Excel!")
        elif not crb_data:
            ng_id.append(f"[NG] ID {feed_id} ada di Excel ('{ex_data['PART']}'), tapi GAADA di Mesin!")
        elif str(ex_data["PART"]).upper() != str(crb_data["PART"]).upper():
            ng_id.append(f"[NG] BEDA PART di ID {feed_id}! Excel='{ex_data['PART']}' vs Mesin='{crb_data['PART']}'")
        elif ex_data["QTY"] != crb_data["QTY"] and not is_table_10_or_11:
            ng_qty.append(f"[NG] BEDA JUMLAH di ID {feed_id}! Excel={ex_data['QTY']} pcs vs Mesin={crb_data['QTY']} pcs")
        else:
            if is_table_10_or_11 and ex_data["QTY"] != crb_data["QTY"]:
                matches.append(
                    f"[MATCH] ID {feed_id} AMAN (Part: {ex_data['PART']}, "
                    f"QTY Beda Diabaikan: Ex={ex_data['QTY']} / Ms={crb_data['QTY']})"
                )
            else:
                matches.append(f"[MATCH] ID {feed_id} AMAN (Part: {ex_data['PART']}, QTY: {ex_data['QTY']} pcs)")

    return WorksheetResult(ng_id, ng_qty, matches)

