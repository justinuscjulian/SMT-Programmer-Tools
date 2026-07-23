import os
import re
from pathlib import Path
from services.feeder_mapping_service import _infer_npm_feeder_id, KNOWN_LINE8_PART_FEEDERS

def pu_code_to_location(pu_code, spans=1):
    try:
        pu = int(pu_code)
    except (ValueError, TypeError):
        return str(pu_code)

    table = pu // 10000
    rem = pu % 10000

    if rem >= 100 or spans == 1:
        if rem >= 100:
            slot = (rem + 1) // 2
            pos = "L" if (rem % 2 == 1) else "R"
            return f"[{table}]{slot:02d}{pos}"
        else:
            if spans > 1:
                return f"[{table}]{rem:02d}-{rem + spans - 1:02d}"
            return f"[{table}]{rem:02d}"
    else:
        if spans > 1:
            return f"[{table}]{rem:02d}-{rem + spans - 1:02d}"
        return f"[{table}]{rem:02d}"

def location_to_pu_code(table, slot, pos=""):
    t = int(table)
    s = int(slot)
    p = str(pos).strip().upper()

    if p == "L":
        rem = (s - 1) * 2 + 1
    elif p == "R":
        rem = (s - 1) * 2 + 2
    else:
        rem = s

    return t * 10000 + rem

class NpmTxtDocument:
    def __init__(self, raw_content=""):
        self.sections = {}
        self.section_order = []
        if raw_content:
            self.parse(raw_content)

    def parse(self, raw_content):
        lines = raw_content.splitlines()
        current_sec = None
        self.sections = {}
        self.section_order = []

        for line in lines:
            s = line.strip()
            if s.startswith('[') and s.endswith(']'):
                current_sec = s
                if current_sec not in self.sections:
                    self.sections[current_sec] = []
                    self.section_order.append(current_sec)
            elif current_sec:
                self.sections[current_sec].append(line)

    def get_parts_map(self):
        parts_by_id = {}
        id_by_part = {}
        sec_name = "[PartsDataEx]" if "[PartsDataEx]" in self.sections else "[PartsData]"
        if sec_name in self.sections:
            for l in self.sections[sec_name]:
                s = l.strip()
                if not s or s.startswith("IDNUM"):
                    continue
                tokens = re.findall(r'"[^"]*"|\S+', s)
                if len(tokens) >= 2:
                    idnum = tokens[0]
                    part_num = tokens[1].replace('"', '')
                    parts_by_id[idnum] = part_num
                    id_by_part[part_num.upper()] = idnum
        return parts_by_id, id_by_part, sec_name

    def get_slots(self):
        parts_by_id, _, _ = self.get_parts_map()
        slots = []

        fixed_lines = self.sections.get("[FixedFeeder]", [])
        for idx, line in enumerate(fixed_lines):
            s = line.strip()
            if not s or s.startswith("IDNUM") or s.startswith("//"):
                continue
            tokens = s.split()
            if len(tokens) >= 14:
                pu_code = tokens[2]
                feeder_id = tokens[3]
                part_id = tokens[13]
                part_num = parts_by_id.get(part_id, f"UNKNOWN_{part_id}")
                
                loc_code = pu_code_to_location(pu_code)
                slots.append({
                    "line_index": idx,
                    "pu_code": pu_code,
                    "location_code": loc_code,
                    "part_number": part_num,
                    "feeder_id": feeder_id,
                    "part_idnum": part_id,
                    "raw_line": line
                })
        return slots

    def add_or_get_part_id(self, part_number):
        part_upper = part_number.strip().upper()
        parts_by_id, id_by_part, sec_name = self.get_parts_map()

        if part_upper in id_by_part:
            return id_by_part[part_upper]

        # Generate new numeric IDNUM
        max_id = 0
        for id_str in parts_by_id.keys():
            try:
                max_id = max(max_id, int(id_str))
            except ValueError:
                pass
        new_idnum = str(max_id + 1)

        new_part_line = f'  {new_idnum} "{part_upper}" "ohm" 1 0 0 0 0 0'
        if sec_name not in self.sections:
            sec_name = "[PartsDataEx]"
            self.sections[sec_name] = ["IDNUM PARTNAME SHAPENAME SPEEDKIND ALIGNKIND GROUPNAME ILLUMKIND MOUNTOPTION ANGLEKIND"]
            if sec_name not in self.section_order:
                self.section_order.append(sec_name)

        self.sections[sec_name].append(new_part_line)
        return new_idnum

    def add_slot(self, table, slot, pos, part_number, feeder_id=None):
        pu_code = str(location_to_pu_code(table, slot, pos))
        part_num = part_number.strip().upper()
        part_idnum = self.add_or_get_part_id(part_num)

        if not feeder_id:
            uses_lr = pos in ("L", "R")
            feeder_id = _infer_npm_feeder_id(part_num, uses_lr_position=uses_lr, spans_slots=1)

        # Check if slot exists in [FixedFeeder]
        fixed_lines = self.sections.get("[FixedFeeder]", [])
        existing_idx = -1
        for idx, line in enumerate(fixed_lines):
            s = line.strip()
            tokens = s.split()
            if len(tokens) >= 3 and tokens[2] == pu_code:
                existing_idx = idx
                break

        fixed_header = "259"
        if fixed_lines:
            for l in fixed_lines:
                toks = l.strip().split()
                if len(toks) > 0 and toks[0].isdigit():
                    fixed_header = toks[0]
                    break

        new_fixed_line = f"  {fixed_header} 0 {pu_code} {feeder_id} 0 0 0 0 0 0 0 0 0 {part_idnum} 0 0 0 0 0 0 0 0 0"
        new_stock_line = f"  109 {pu_code} {part_idnum} 0 0 0 0 0 0 0 0 0 {feeder_id} 0 0 0 0 0 0 0 0 0 0 0.000 0"

        if "[FixedFeeder]" not in self.sections:
            self.sections["[FixedFeeder]"] = []
            if "[FixedFeeder]" not in self.section_order:
                self.section_order.append("[FixedFeeder]")

        if existing_idx >= 0:
            self.sections["[FixedFeeder]"][existing_idx] = new_fixed_line
        else:
            self.sections["[FixedFeeder]"].append(new_fixed_line)

        # Update [StockData]
        if "[StockData]" not in self.sections:
            self.sections["[StockData]"] = []
            if "[StockData]" not in self.section_order:
                self.section_order.append("[StockData]")

        stock_lines = self.sections["[StockData]"]
        existing_stock_idx = -1
        for idx, line in enumerate(stock_lines):
            s = line.strip()
            tokens = s.split()
            if len(tokens) >= 2 and tokens[1] == pu_code:
                existing_stock_idx = idx
                break

        if existing_stock_idx >= 0:
            self.sections["[StockData]"][existing_stock_idx] = new_stock_line
        else:
            self.sections["[StockData]"].append(new_stock_line)

    def delete_slot(self, pu_code):
        pu_str = str(pu_code)
        if "[FixedFeeder]" in self.sections:
            self.sections["[FixedFeeder]"] = [
                l for l in self.sections["[FixedFeeder]"]
                if not (len(l.strip().split()) >= 3 and l.strip().split()[2] == pu_str)
            ]
        if "[StockData]" in self.sections:
            self.sections["[StockData]"] = [
                l for l in self.sections["[StockData]"]
                if not (len(l.strip().split()) >= 2 and l.strip().split()[1] == pu_str)
            ]

    def serialize(self):
        output_lines = []
        for sec in self.section_order:
            output_lines.append(sec)
            for line in self.sections[sec]:
                output_lines.append(line)
        return "\n".join(output_lines) + "\n"

def load_npm_txt(file_path):
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
    doc = NpmTxtDocument(content)
    return doc

def save_npm_txt(doc, output_path):
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    content = doc.serialize()
    with open(out, "w", encoding="utf-8") as f:
        f.write(content)
