import csv
import os
import re
import shutil
from collections import Counter, OrderedDict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from services.errors import ServiceError


CRB_EXTENSION = ".crb"
LINE_PATTERN = re.compile(r"(?<![A-Za-z0-9])INI\d+(?![A-Za-z0-9])", re.IGNORECASE)


@dataclass
class CrbProgramCollectorConfig:
    source_folder: str
    destination_folder: str
    part_numbers_text: str
    line_filter_text: str = ""


@dataclass
class CrbProgramMatch:
    source_path: str
    source_folder: str
    file_name: str
    size: int
    matched_parts: list[str]
    line: str


@dataclass
class CrbProgramScanResult:
    source_folder: str
    destination_folder: str
    part_numbers: list[str]
    line_filters: list[str]
    matches: list[CrbProgramMatch]
    total_crb_scanned: int
    walk_errors: list[str] = field(default_factory=list)
    part_pre_line_counts: dict[str, int] = field(default_factory=dict)
    part_match_counts: dict[str, int] = field(default_factory=dict)
    line_counts: dict[str, int] = field(default_factory=dict)

    @property
    def found_parts(self):
        return [part for part in self.part_numbers if self.part_match_counts.get(part, 0) > 0]

    @property
    def missing_parts(self):
        return [part for part in self.part_numbers if self.part_pre_line_counts.get(part, 0) == 0]

    @property
    def line_blocked_parts(self):
        return [
            part
            for part in self.part_numbers
            if self.part_pre_line_counts.get(part, 0) > 0 and self.part_match_counts.get(part, 0) == 0
        ]


@dataclass
class CrbProgramCopyRow:
    source_path: str
    destination_path: str
    file_name: str
    matched_parts: list[str]
    status: str
    message: str
    source_size: int = 0
    destination_size: int = 0


@dataclass
class CrbProgramCopyResult:
    scan_result: CrbProgramScanResult
    rows: list[CrbProgramCopyRow]
    copied_count: int
    error_count: int
    destination_folder: str


def scan_crb_programs(config: CrbProgramCollectorConfig, progress_callback=None):
    source_folder, part_numbers, line_filters = _validate_scan_config(config)
    destination_folder = str(Path(config.destination_folder).resolve()) if str(config.destination_folder or "").strip() else ""

    _emit_progress(progress_callback, 0, "Scanning .crb files...")
    crb_files, walk_errors = _find_crb_files(source_folder)
    total_files = len(crb_files)

    pre_line_counts = {part: 0 for part in part_numbers}
    match_counts = {part: 0 for part in part_numbers}
    line_counts = Counter()
    matches = []

    if not crb_files:
        _emit_progress(progress_callback, 100, "No .crb files found")
        return CrbProgramScanResult(
            source_folder=str(source_folder),
            destination_folder=destination_folder,
            part_numbers=part_numbers,
            line_filters=line_filters,
            matches=[],
            total_crb_scanned=0,
            walk_errors=walk_errors,
            part_pre_line_counts=pre_line_counts,
            part_match_counts=match_counts,
            line_counts={},
        )

    for index, file_path in enumerate(crb_files, start=1):
        percent = max(1, min(99, int((index - 1) / total_files * 100)))
        _emit_progress(progress_callback, percent, f"Scanning {index}/{total_files}: {file_path.name}")

        matched_parts = _matched_parts_for_path(file_path, part_numbers)
        if not matched_parts:
            continue

        for part in matched_parts:
            pre_line_counts[part] += 1

        if not _line_filter_matches(file_path.name, line_filters):
            continue

        line = _extract_line_indicator(file_path.name)
        file_size = _safe_file_size(file_path)
        matches.append(
            CrbProgramMatch(
                source_path=str(file_path),
                source_folder=str(file_path.parent),
                file_name=file_path.name,
                size=file_size,
                matched_parts=matched_parts,
                line=line,
            )
        )
        line_counts[line or "UNKNOWN"] += 1
        for part in matched_parts:
            match_counts[part] += 1

    _emit_progress(progress_callback, 100, f"Scan complete: {len(matches)} file(s) matched")
    return CrbProgramScanResult(
        source_folder=str(source_folder),
        destination_folder=destination_folder,
        part_numbers=part_numbers,
        line_filters=line_filters,
        matches=matches,
        total_crb_scanned=total_files,
        walk_errors=walk_errors,
        part_pre_line_counts=pre_line_counts,
        part_match_counts=match_counts,
        line_counts=dict(sorted(line_counts.items(), key=lambda item: _line_sort_key(item[0]))),
    )


def copy_crb_programs(scan_result: CrbProgramScanResult, destination_folder=None, progress_callback=None):
    if scan_result is None:
        raise ServiceError("Belum ada hasil scan untuk dicopy.", title="Data kosong")
    if not scan_result.matches:
        raise ServiceError("Tidak ada file match untuk dicopy.", title="Data kosong")

    destination_text = str(destination_folder or scan_result.destination_folder or "").strip()
    if not destination_text:
        raise ServiceError("Folder tujuan belum dipilih.", title="Input belum lengkap")
    destination = Path(destination_text).expanduser()

    try:
        destination.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise ServiceError(f"Folder tujuan tidak bisa dibuat/diakses:\n{destination}\n{exc}", title="Permission error") from exc

    name_counts = Counter(match.file_name.casefold() for match in scan_result.matches)
    reserved_paths = set()
    rows = []
    total_matches = len(scan_result.matches)

    for index, match in enumerate(scan_result.matches, start=1):
        source = Path(match.source_path)
        percent = max(1, min(99, int((index - 1) / total_matches * 100)))
        _emit_progress(progress_callback, percent, f"Copying {index}/{total_matches}: {source.name}")

        destination_path = _unique_destination_path(
            source,
            destination,
            reserved_paths,
            force_suffix=name_counts[match.file_name.casefold()] > 1,
        )
        reserved_paths.add(str(destination_path).casefold())

        try:
            shutil.copy2(source, destination_path)
            source_size = source.stat().st_size
            destination_size = destination_path.stat().st_size
            if source_size != destination_size:
                rows.append(
                    CrbProgramCopyRow(
                        source_path=str(source),
                        destination_path=str(destination_path),
                        file_name=source.name,
                        matched_parts=match.matched_parts,
                        status="VERIFY_FAILED",
                        message=f"Size mismatch: source {source_size} bytes, destination {destination_size} bytes",
                        source_size=source_size,
                        destination_size=destination_size,
                    )
                )
                continue

            rows.append(
                CrbProgramCopyRow(
                    source_path=str(source),
                    destination_path=str(destination_path),
                    file_name=source.name,
                    matched_parts=match.matched_parts,
                    status="COPIED",
                    message="OK",
                    source_size=source_size,
                    destination_size=destination_size,
                )
            )
        except OSError as exc:
            rows.append(
                CrbProgramCopyRow(
                    source_path=str(source),
                    destination_path=str(destination_path),
                    file_name=source.name,
                    matched_parts=match.matched_parts,
                    status="ERROR",
                    message=str(exc),
                )
            )

    copied_count = sum(1 for row in rows if row.status == "COPIED")
    error_count = len(rows) - copied_count
    _emit_progress(progress_callback, 100, f"Copy complete: {copied_count}/{total_matches} copied")
    return CrbProgramCopyResult(
        scan_result=scan_result,
        rows=rows,
        copied_count=copied_count,
        error_count=error_count,
        destination_folder=str(destination),
    )


def suggest_report_name():
    return f"CRB_Program_Collector_Report_{datetime.now().strftime('%y%m%d_%H%M')}.txt"


def export_crb_program_report(scan_result, output_path, copy_result=None):
    if scan_result is None:
        raise ServiceError("Belum ada hasil scan untuk diexport.", title="Data kosong")

    output = Path(output_path)
    if output.suffix.lower() not in {".txt", ".csv"}:
        output = output.with_suffix(".txt")
    output.parent.mkdir(parents=True, exist_ok=True)

    if output.suffix.lower() == ".csv":
        _export_csv_report(scan_result, output, copy_result)
    else:
        output.write_text(format_crb_program_report(scan_result, copy_result), encoding="utf-8-sig")
    return str(output)


def format_crb_program_report(scan_result, copy_result=None):
    if scan_result is None:
        return ""

    copied_count = copy_result.copied_count if copy_result else 0
    error_count = copy_result.error_count if copy_result else 0
    line_filter_text = ", ".join(scan_result.line_filters) if scan_result.line_filters else "ALL"

    lines = [
        "CRB Program Finder/Collector Report",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        f"Source folder: {scan_result.source_folder}",
        f"Destination folder: {(copy_result.destination_folder if copy_result else scan_result.destination_folder) or '-'}",
        f"PCB part numbers: {', '.join(scan_result.part_numbers)}",
        f"Line filter: {line_filter_text}",
        "",
        f"Total .crb scanned: {scan_result.total_crb_scanned}",
        f"Total file match: {len(scan_result.matches)}",
        f"Total file copied: {copied_count}",
        f"Total error/verify failed: {error_count}",
        f"Folder scan errors: {len(scan_result.walk_errors)}",
        "",
        "Found part numbers:",
        *_prefixed(scan_result.found_parts or ["-"]),
        "",
        "Part found, but no matching file for selected line:",
        *_prefixed(scan_result.line_blocked_parts or ["-"]),
        "",
        "Part numbers not found:",
        *_prefixed(scan_result.missing_parts or ["-"]),
        "",
        "Count file per part number:",
        *_prefixed([f"{part}: {scan_result.part_match_counts.get(part, 0)} files" for part in scan_result.part_numbers]),
        "",
        "Count file per line:",
        *_prefixed([f"{line}: {count} files" for line, count in scan_result.line_counts.items()] or ["-"]),
        "",
        "Matched files:",
    ]

    for match in scan_result.matches:
        lines.append(f"- {match.file_name} | Line: {match.line or '-'} | Part: {', '.join(match.matched_parts)} | {match.source_path}")

    if copy_result:
        lines.append("")
        lines.append("Copy result:")
        for row in copy_result.rows:
            lines.append(
                f"- {row.status}: {row.file_name} -> {row.destination_path}"
                + (f" | {row.message}" if row.message and row.message != "OK" else "")
            )

    if scan_result.walk_errors:
        lines.append("")
        lines.append("Folder scan errors:")
        lines.extend(_prefixed(scan_result.walk_errors))

    return "\n".join(lines)


def _export_csv_report(scan_result, output_path, copy_result=None):
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Section", "Key", "Value"])
        writer.writerow(["Summary", "Source folder", scan_result.source_folder])
        writer.writerow(["Summary", "Destination folder", (copy_result.destination_folder if copy_result else scan_result.destination_folder) or ""])
        writer.writerow(["Summary", "Line filter", ", ".join(scan_result.line_filters) if scan_result.line_filters else "ALL"])
        writer.writerow(["Summary", "Total .crb scanned", scan_result.total_crb_scanned])
        writer.writerow(["Summary", "Total file match", len(scan_result.matches)])
        writer.writerow(["Summary", "Total file copied", copy_result.copied_count if copy_result else 0])
        writer.writerow(["Summary", "Total error/verify failed", copy_result.error_count if copy_result else 0])
        writer.writerow([])
        writer.writerow(["Matched Files"])
        writer.writerow(["File Name", "Line", "Matched Part Numbers", "Source Path", "Size"])
        for match in scan_result.matches:
            writer.writerow([match.file_name, match.line, ", ".join(match.matched_parts), match.source_path, match.size])
        writer.writerow([])
        writer.writerow(["Part Counts"])
        writer.writerow(["Part Number", "Matched Files", "Found Before Line Filter"])
        for part in scan_result.part_numbers:
            writer.writerow([part, scan_result.part_match_counts.get(part, 0), scan_result.part_pre_line_counts.get(part, 0)])
        writer.writerow([])
        writer.writerow(["Line Counts"])
        writer.writerow(["Line", "Matched Files"])
        for line, count in scan_result.line_counts.items():
            writer.writerow([line, count])
        if copy_result:
            writer.writerow([])
            writer.writerow(["Copy Result"])
            writer.writerow(["Status", "File Name", "Matched Part Numbers", "Destination Path", "Message"])
            for row in copy_result.rows:
                writer.writerow([row.status, row.file_name, ", ".join(row.matched_parts), row.destination_path, row.message])


def parse_part_numbers(value):
    unique = OrderedDict()
    for token in re.split(r"[\r\n,;]+", str(value or "")):
        clean = token.strip().upper()
        if clean:
            unique.setdefault(clean, clean)
    return list(unique.values())


def parse_line_filters(value):
    unique = OrderedDict()
    for token in re.split(r"[\s,;]+", str(value or "")):
        clean = token.strip().upper()
        if not clean:
            continue
        if clean.isdigit():
            clean = f"INI{clean}"
        unique.setdefault(clean, clean)
    return list(unique.values())


def _validate_scan_config(config):
    source_text = str(config.source_folder or "").strip()
    if not source_text:
        raise ServiceError("Folder sumber belum dipilih.", title="Input belum lengkap")
    source_folder = Path(source_text).expanduser()
    if not source_folder.is_dir():
        raise ServiceError(f"Folder sumber tidak ditemukan:\n{source_folder}", title="Folder tidak ditemukan")

    part_numbers = parse_part_numbers(config.part_numbers_text)
    if not part_numbers:
        raise ServiceError("Daftar PCB part number belum diisi.", title="Input belum lengkap")

    return source_folder, part_numbers, parse_line_filters(config.line_filter_text)


def _find_crb_files(source_folder):
    files = []
    errors = []

    def on_error(error):
        filename = getattr(error, "filename", "")
        message = getattr(error, "strerror", str(error))
        errors.append(f"{filename}: {message}")

    for current_folder, dir_names, file_names in os.walk(source_folder, onerror=on_error):
        dir_names.sort(key=lambda name: name.casefold())
        for file_name in sorted(file_names, key=lambda name: name.casefold()):
            path = Path(current_folder) / file_name
            if path.suffix.casefold() == CRB_EXTENSION:
                files.append(path)

    return files, errors


def _matched_parts_for_path(file_path, part_numbers):
    text = str(file_path).casefold()
    return [part for part in part_numbers if part.casefold() in text]


def _line_filter_matches(file_name, line_filters):
    if not line_filters:
        return True
    for line_filter in line_filters:
        pattern = re.compile(rf"(?<![A-Za-z0-9]){re.escape(line_filter)}(?![A-Za-z0-9])", re.IGNORECASE)
        if pattern.search(file_name):
            return True
    return False


def _extract_line_indicator(file_name):
    match = LINE_PATTERN.search(str(file_name or ""))
    return match.group(0).upper() if match else ""


def _safe_file_size(file_path):
    try:
        return Path(file_path).stat().st_size
    except OSError:
        return 0


def _unique_destination_path(source_path, destination_folder, reserved_paths, force_suffix=False):
    source_path = Path(source_path)
    destination_folder = Path(destination_folder)
    candidate = destination_folder / source_path.name
    if not force_suffix and not candidate.exists() and str(candidate).casefold() not in reserved_paths:
        return candidate

    suffix = _safe_filename_part(source_path.parent.name) or "source"
    base_name = f"{source_path.stem}__from_{suffix}"
    candidate = destination_folder / f"{base_name}{source_path.suffix}"
    if not candidate.exists() and str(candidate).casefold() not in reserved_paths:
        return candidate

    counter = 2
    while True:
        candidate = destination_folder / f"{base_name}_{counter}{source_path.suffix}"
        if not candidate.exists() and str(candidate).casefold() not in reserved_paths:
            return candidate
        counter += 1


def _safe_filename_part(value):
    text = re.sub(r'[<>:"/\\|?*\r\n]+', "_", str(value or "")).strip(" ._")
    return text[:80]


def _prefixed(values):
    return [f"- {value}" for value in values]


def _line_sort_key(value):
    match = re.search(r"(\d+)", str(value or ""))
    if match:
        return (0, int(match.group(1)), str(value).upper())
    return (1, 0, str(value).upper())


def _emit_progress(progress_callback, percent, message):
    if progress_callback:
        progress_callback(percent, message)
