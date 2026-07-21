import re
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from services.errors import ServiceError
from services.model_feeder_group_service import _scan_models, _emit_progress, ModelFeederGroupConfig, _build_pair_rows, _build_groups

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

class VirtualMachine:
    def __init__(self, line_type="Line 1-5"):
        self.tables = {}
        if line_type == "Line 1-5":
            for t in range(1, 5):
                self.tables[t] = [{'L': False, 'R': False} for _ in range(17)]
            for t in range(5, 9):
                self.tables[t] = [{'L': False, 'R': False} for _ in range(30)]
            self.tables[9] = [{'L': False, 'R': False} for _ in range(30)]
        elif line_type == "Line 8":
            for t in range(1, 7):
                self.tables[t] = [{'L': False, 'R': False} for _ in range(30)]
            self.tables[7] = [{'L': False, 'R': False} for _ in range(30)]
        else: # Line 6-7
            for t in range(1, 10):
                self.tables[t] = [{'L': False, 'R': False} for _ in range(30)]

    def _parse_loc(self, loc_str):
        match = re.match(r'^\[(\d+)\](\d+)(?:-(\d+))?([LRlr])?$', loc_str)
        if not match:
            return None
        t = int(match.group(1))
        s1 = int(match.group(2))
        s2 = int(match.group(3)) if match.group(3) else s1
        pos = match.group(4).upper() if match.group(4) else None
        return t, s1, s2, pos

    def can_add(self, loc_str):
        parsed = self._parse_loc(loc_str)
        if not parsed:
            return False
        t, s1, s2, pos = parsed
        
        if t not in self.tables:
            return False
            
        if s1 < 1 or s2 > len(self.tables[t]):
            return False
            
        for s in range(s1, s2 + 1):
            idx = s - 1
            if pos == 'L':
                if self.tables[t][idx]['L']: return False
            elif pos == 'R':
                if self.tables[t][idx]['R']: return False
            else:
                if self.tables[t][idx]['L'] or self.tables[t][idx]['R']: return False
                
        return True

    def add(self, loc_str):
        if not self.can_add(loc_str):
            raise ValueError(f'Cannot add {loc_str}')
            
        t, s1, s2, pos = self._parse_loc(loc_str)
        for s in range(s1, s2 + 1):
            idx = s - 1
            if pos == 'L':
                self.tables[t][idx]['L'] = True
            elif pos == 'R':
                self.tables[t][idx]['R'] = True
            else:
                self.tables[t][idx]['L'] = True
                self.tables[t][idx]['R'] = True

    def copy(self):
        import copy
        new_vm = VirtualMachine()
        new_vm.tables = copy.deepcopy(self.tables)
        return new_vm

    def find_fallback(self, loc_str):
        parsed = self._parse_loc(loc_str)
        if not parsed:
            return None
        pref_t, s1, s2, pos = parsed
        needs_slots = s2 - s1 + 1
        
        if pref_t not in self.tables:
            return None
            
        allowed_tables = [pref_t]
            
        for t in allowed_tables:
            if t not in self.tables:
                continue
                
            num_slots = self.tables[t]
            for start_s in range(1, len(num_slots) - needs_slots + 2):
                end_s = start_s + needs_slots - 1
                idx = start_s - 1
                
                if pos in ['L', 'R'] and needs_slots == 1:
                    # Can fallback to either L or R on this slot
                    if not num_slots[idx]['L']:
                        return f"[{t}]{start_s}L"
                    if not num_slots[idx]['R']:
                        return f"[{t}]{start_s}R"
                else:
                    # Needs full slot(s)
                    can_fit = True
                    for s in range(start_s, end_s + 1):
                        s_idx = s - 1
                        if num_slots[s_idx]['L'] or num_slots[s_idx]['R']:
                            can_fit = False
                            break
                    if can_fit:
                        if needs_slots > 1:
                            return f"[{t}]{start_s}-{end_s}"
                        else:
                            return f"[{t}]{start_s}"
        return None


class PcbInfo:
    def __init__(self, part_number, file_path, components, insert_averages=None):
        self.part_number = part_number
        self.file_path = file_path
        self.components = components
        self.insert_averages = insert_averages or {}


class GroupResult:
    def __init__(self, group_name, pcbs, slot_mapping, part_mapping, unassigned_parts=None):
        self.group_name = group_name
        self.pcbs = pcbs
        self.slot_mapping = slot_mapping
        self.part_mapping = part_mapping
        self.unassigned_parts = unassigned_parts or []

def get_master_mapping(excel_path, line_type=None):
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active
    
    headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
    try:
        idx_part = headers.index("Part Number")
        idx_loc = headers.index("Feeder Paling Sering")
        idx_freq = headers.index("Total Muncul")
        idx_other = headers.index("Feeder Lain") if "Feeder Lain" in headers else -1
    except ValueError as e:
        raise ServiceError(f"Format Master Mapping salah: {e}")
        
    master = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[idx_part]:
            continue
        part = str(row[idx_part]).strip().upper()
        loc = str(row[idx_loc]).strip() if row[idx_loc] else ""
        freq = int(row[idx_freq]) if row[idx_freq] else 0
        
        # Ignore Tables
        if "[10]" in loc:
            continue
        if line_type == "Line 8" and ("[8]" in loc or "[9]" in loc):
            continue
            
        other_locs = []
        if idx_other != -1 and row[idx_other]:
            matches = re.findall(r'(\[\d+\]\d+(?:-\d+)?(?:[LRlr])?)', str(row[idx_other]))
            for m in matches:
                if "[10]" in m:
                    continue
                if line_type == "Line 8" and ("[8]" in m or "[9]" in m):
                    continue
                other_locs.append(m)

        if part not in master:
            master[part] = []
        master[part].append({
            "location": loc,
            "frequency": freq,
            "alternatives": other_locs
        })
    return master

def generate_all_table_groups(crb_folder, master_excel_path, target_pcbs_text, line_type, min_sim, min_shared, progress_callback=None):
    _emit_progress(progress_callback, 0, "Membaca referensi & Master Mapping...")
    master = get_master_mapping(master_excel_path, line_type)
    
    # Strip L/R suffixes for tables that are for large/full-slot components
    new_master = {}
    for part, loc_list in master.items():
        new_items = []
        for item in loc_list:
            loc = item["location"]
            if not loc:
                continue
            
            if line_type == "Line 8":
                # For Line 8, Table 5 and Table 7 are for large components (no L/R)
                if "[5]" in loc or "[7]" in loc:
                    match = re.match(r'^(\[(?:5|7)\]\d+(?:-\d+)?)[LRlr]?$', loc)
                    if match:
                        loc = match.group(1)
            else:
                # For Line 1-5 and Line 6-7, Table 9 is for large components (no L/R)
                if "[9]" in loc:
                    match = re.match(r'^(\[9\]\d+(?:-\d+)?)[LRlr]?$', loc)
                    if match:
                        loc = match.group(1)
            
            # Clean alternatives similarly
            new_alts = []
            for alt in item.get("alternatives", []):
                if line_type == "Line 8":
                    if "[5]" in alt or "[7]" in alt:
                        match = re.match(r'^(\[(?:5|7)\]\d+(?:-\d+)?)[LRlr]?$', alt)
                        if match:
                            alt = match.group(1)
                else:
                    if "[9]" in alt:
                        match = re.match(r'^(\[9\]\d+(?:-\d+)?)[LRlr]?$', alt)
                        if match:
                            alt = match.group(1)
                new_alts.append(alt)
            
            new_item = dict(item)
            new_item["location"] = loc
            new_item["alternatives"] = new_alts
            new_items.append(new_item)
        if new_items:
            new_master[part] = new_items
    master = new_master
    
    target_list = []
    if target_pcbs_text:
        for line in target_pcbs_text.splitlines():
            line = line.strip().upper()
            if line:
                target_list.append(line)

    _emit_progress(progress_callback, 10, "Scanning PCB folders...")
    models, _, _, _ = _scan_models(crb_folder, target_list, progress_callback)
    if not models:
        raise ServiceError("Tidak ada file Excel program yang valid ditemukan atau sesuai dengan target PCB.")

    # Step 1: Group PCBs based on similarity
    _emit_progress(progress_callback, 40, "Mengelompokkan PCB berdasarkan kemiripan...")
    config = ModelFeederGroupConfig(
        source_folder=crb_folder,
        min_similarity_percent=min_sim,
        min_shared_components=min_shared,
        target_pcb_list=target_list
    )
    pair_rows, pair_lookup = _build_pair_rows(models, config)
    raw_groups = _build_groups(models, pair_rows, pair_lookup)
    
    groups = []
    _emit_progress(progress_callback, 60, "Membangun setup feeder untuk masing-masing grup...")
    
    total_groups = len(raw_groups)
    
    for i, raw_group in enumerate(raw_groups):
        group_pcbs = []
        for pcb_name in raw_group:
            model = models[pcb_name]
            group_pcbs.append(PcbInfo(
                part_number=model.display_name, # clean name instead of path
                file_path="",
                components=set(str(val).strip().upper() for val in model.components.values() if val),
                insert_averages=model.insert_averages
            ))
        
        # Calculate component frequency strictly within this group
        group_part_freq = {}
        group_all_parts = set()
        for pcb in group_pcbs:
            for part in pcb.components:
                group_part_freq[part] = group_part_freq.get(part, 0) + 1
                group_all_parts.add(part)
                
        # Sort components: primarily by frequency in this group, then by global master frequency
        def part_sort_key(p):
            local_freq = group_part_freq.get(p, 0)
            loc_list = master.get(p, [])
            master_freq = max([item["frequency"] for item in loc_list], default=0)
            return (local_freq, master_freq)
            
        sorted_parts = sorted(list(group_all_parts), key=part_sort_key, reverse=True)
        
        vm = VirtualMachine(line_type=line_type)
        slot_mapping = {}
        part_mapping = {}
        unassigned_parts = []
        
        for part in sorted_parts:
            loc_list = master.get(part, [])
            if not loc_list:
                unassigned_parts.append(part)
                continue
                
            loc_list = sorted(loc_list, key=lambda x: x["frequency"], reverse=True)
            
            # Check avg inserts in this group
            inserts_in_group = [pcb.insert_averages.get(part, 0) for pcb in group_pcbs]
            avg_inserts = sum(inserts_in_group) / len(group_pcbs) if group_pcbs else 0
            
            is_balancing = (avg_inserts >= 20) and (len(loc_list) > 1)
            
            if is_balancing:
                placed_slots = []
                for loc_item in loc_list:
                    loc = loc_item["location"]
                    if not loc:
                        continue
                    if vm.can_add(loc):
                        vm.add(loc)
                        slot_mapping[loc] = part
                        placed_slots.append(loc)
                    else:
                        fallback = vm.find_fallback(loc)
                        if fallback:
                            vm.add(fallback)
                            slot_mapping[fallback] = part
                            placed_slots.append(fallback)
                if placed_slots:
                    part_mapping[part] = placed_slots[0]
                else:
                    unassigned_parts.append(part)
            else:
                primary_loc = loc_list[0]["location"]
                if not primary_loc:
                    unassigned_parts.append(part)
                    continue
                
                # Build candidates list: primary location first, then alternatives in order, removing duplicates
                candidates = [primary_loc]
                for alt in loc_list[0].get("alternatives", []):
                    if alt and alt not in candidates:
                        candidates.append(alt)
                
                placed = False
                for loc in candidates:
                    if vm.can_add(loc):
                        vm.add(loc)
                        slot_mapping[loc] = part
                        part_mapping[part] = loc
                        placed = True
                        break
                    else:
                        fallback = vm.find_fallback(loc)
                        if fallback:
                            vm.add(fallback)
                            slot_mapping[fallback] = part
                            part_mapping[part] = fallback
                            placed = True
                            break
                            
                if not placed:
                    unassigned_parts.append(part)
                    
        groups.append(GroupResult(
            group_name=f"Group {i + 1}",
            pcbs=group_pcbs,
            slot_mapping=slot_mapping,
            part_mapping=part_mapping,
            unassigned_parts=unassigned_parts
        ))
        
        percent = 60 + int((i + 1) / max(1, total_groups) * 35)
        _emit_progress(progress_callback, percent, f"Setup Group {i + 1} selesai...")
        
    _emit_progress(progress_callback, 100, "Selesai")
    return groups

def export_all_table_groups(groups, output_path):
    wb = Workbook()
    
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.append(["Group Name", "PCB Name", "Total PCBs in Group", "Total Fixed Slots", "Total Skipped Components"])
    
    for g in groups:
        for p in g.pcbs:
            row = [g.group_name, p.part_number, len(g.pcbs), len(g.slot_mapping), len(g.unassigned_parts)]
            ws_summary.append(row)
        
    # Calculate base components used in ALL groups
    group_parts = []
    for g in groups:
        parts_in_group = set(g.slot_mapping.values())
        group_parts.append(parts_in_group)
        
    common_parts = set()
    if group_parts:
        common_parts = group_parts[0]
        for s in group_parts[1:]:
            common_parts = common_parts & s
            
    common_parts_list = sorted(list(common_parts))
    
    # Style and write base components
    ws_summary.append([])
    ws_summary.append([])
    
    title_font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    ws_summary.append(["Base Components (Kerangka Dasar - Terpakai di Semua Group)"])
    title_row_idx = ws_summary.max_row
    ws_summary.cell(row=title_row_idx, column=1).font = title_font
    
    base_headers = ["Base Component P/N"] + [f"{g.group_name} Slot" for g in groups]
    ws_summary.append(base_headers)
    header_row_idx = ws_summary.max_row
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    
    for col_idx in range(1, len(base_headers) + 1):
        cell = ws_summary.cell(row=header_row_idx, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align if col_idx > 1 else left_align
        
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )
    
    body_font = Font(name="Calibri", size=11)
    
    if common_parts_list:
        for part in common_parts_list:
            row = [part]
            for g in groups:
                slots = [loc for loc, p in g.slot_mapping.items() if p == part]
                row.append(", ".join(slots))
            ws_summary.append(row)
            
            data_row_idx = ws_summary.max_row
            for col_idx in range(1, len(row) + 1):
                cell = ws_summary.cell(row=data_row_idx, column=col_idx)
                cell.font = body_font
                cell.alignment = center_align if col_idx > 1 else left_align
                cell.border = thin_border
    else:
        ws_summary.append(["Tidak ada komponen yang terpakai di semua group."])
        data_row_idx = ws_summary.max_row
        cell = ws_summary.cell(row=data_row_idx, column=1)
        cell.font = body_font
        cell.border = thin_border
        
    # Auto-adjust column widths on summary sheet
    for col in ws_summary.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value or "")
            # Skip checking the title row since it's long and would stretch column A too much
            if cell.row == title_row_idx:
                continue
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws_summary.column_dimensions[col_letter].width = max(max_len + 3, 15)
        
    for g in groups:
        ws = wb.create_sheet(title=g.group_name)
        ws.append(["Table", "Slot", "Position", "Location Code", "Part Number"])
        
        records = []
        for loc, part in g.slot_mapping.items():
            match = re.match(r"^\[(\d+)\](\d+)(?:-\d+)?([A-Za-z]+)?$", loc)
            if match:
                table = int(match.group(1))
                slot = int(match.group(2))
                pos = str(match.group(3) or "").upper()
                records.append({
                    "table": table,
                    "slot": slot,
                    "pos": pos,
                    "loc": loc,
                    "part": part
                })
            else:
                records.append({
                    "table": 99,
                    "slot": 99,
                    "pos": "",
                    "loc": loc,
                    "part": part
                })
                
        records.sort(key=lambda x: (x["table"], x["slot"], x["pos"], natural_sort_key(x["loc"])))
        
        for r in records:
            ws.append([r["table"], r["slot"], r["pos"], r["loc"], r["part"]])
            
        # Write skipped parts at the bottom
        if g.unassigned_parts:
            ws.append([])
            ws.append(["SKIPPED / DYNAMIC COMPONENTS (Not enough slots or not in master):"])
            ws.append(["Part Number"])
            for p in g.unassigned_parts:
                ws.append([p])
            
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 25
        
    wb.save(output_path)
    return output_path
