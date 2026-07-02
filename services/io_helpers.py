import os
from pathlib import Path


def excel_engine_for_path(path):
    suffix = Path(path).suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return "openpyxl"
    if suffix == ".xls":
        return "xlrd"
    return None


def open_pandas_excel_file(pd_module, path):
    engine = excel_engine_for_path(path)
    if engine:
        return pd_module.ExcelFile(path, engine=engine)
    return pd_module.ExcelFile(path)


def scan_direct_files(folder, extensions, skip_prefixes=(), sort_key=None):
    folder = Path(folder)
    extensions = _normalize_extensions(extensions)
    sort_key = sort_key or (lambda name: name.upper())
    files = []

    with os.scandir(folder) as entries:
        entry_list = sorted(entries, key=lambda entry: sort_key(entry.name))

    for entry in entry_list:
        if not _is_file(entry):
            continue
        if _has_skipped_prefix(entry.name, skip_prefixes):
            continue
        path = Path(entry.path)
        if path.suffix.lower() in extensions:
            files.append(path)
    return files


def scan_recursive_files(source_folder, extensions, skip_prefixes=(), sort_key=None):
    source_folder = Path(source_folder)
    extensions = _normalize_extensions(extensions)
    sort_key = sort_key or (lambda name: name.upper())
    files = []
    errors = []

    def scan_folder(folder):
        try:
            with os.scandir(folder) as entries:
                entry_list = list(entries)
        except OSError as exc:
            errors.append(_format_scan_error(exc, folder))
            return

        dir_entries = []
        file_entries = []
        for entry in entry_list:
            try:
                if entry.is_dir(follow_symlinks=False):
                    dir_entries.append(entry)
                elif entry.is_file():
                    file_entries.append(entry)
            except OSError as exc:
                errors.append(_format_scan_error(exc, entry.path))

        dir_entries.sort(key=lambda entry: sort_key(entry.name))
        file_entries.sort(key=lambda entry: sort_key(entry.name))

        for entry in file_entries:
            if _has_skipped_prefix(entry.name, skip_prefixes):
                continue
            path = Path(entry.path)
            if path.suffix.lower() in extensions:
                files.append(path)

        for entry in dir_entries:
            scan_folder(Path(entry.path))

    scan_folder(source_folder)
    return files, errors


def scan_recursive_dirs(source_folder, include_root=False, sort_key=None):
    source_folder = Path(source_folder)
    sort_key = sort_key or (lambda name: name.upper())
    folders = [source_folder] if include_root else []
    errors = []

    def scan_folder(folder):
        try:
            with os.scandir(folder) as entries:
                dir_entries = [
                    entry
                    for entry in entries
                    if _is_dir(entry)
                ]
        except OSError as exc:
            errors.append(_format_scan_error(exc, folder))
            return

        dir_entries.sort(key=lambda entry: sort_key(entry.name))
        folders.extend(Path(entry.path) for entry in dir_entries)
        for entry in dir_entries:
            scan_folder(Path(entry.path))

    scan_folder(source_folder)
    return folders, errors


def _normalize_extensions(extensions):
    return {str(extension).lower() for extension in extensions}


def _has_skipped_prefix(name, skip_prefixes):
    return any(name.startswith(prefix) for prefix in skip_prefixes)


def _is_file(entry):
    try:
        return entry.is_file()
    except OSError:
        return False


def _is_dir(entry):
    try:
        return entry.is_dir(follow_symlinks=False)
    except OSError:
        return False


def _format_scan_error(error, fallback_path):
    filename = getattr(error, "filename", None) or str(fallback_path)
    message = getattr(error, "strerror", None) or str(error)
    return f"{filename}: {message}"
