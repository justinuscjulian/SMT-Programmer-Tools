import json
from pathlib import Path

import pandas as pd

from services.errors import ServiceError
from utils.paths import app_data_path, legacy_data_path


HISTORY_FILE = "history.json"
MAX_HISTORY = 50


def _active_history_path():
    current_path = app_data_path(HISTORY_FILE)
    if current_path.exists():
        return current_path

    legacy_path = legacy_data_path(HISTORY_FILE)
    if legacy_path.exists():
        return legacy_path

    return current_path


def _write_history_path():
    return app_data_path(HISTORY_FILE)


def load_history():
    history_path = _active_history_path()
    if not history_path.exists():
        return []

    try:
        with Path(history_path).open("r", encoding="utf-8") as handle:
            return json.load(handle)
    except Exception:
        return []


def save_history(entry):
    history = load_history()
    history.insert(0, entry)
    history = history[:MAX_HISTORY]

    try:
        path = _write_history_path()
        with path.open("w", encoding="utf-8") as handle:
            json.dump(history, handle, indent=4)
    except Exception as exc:
        raise ServiceError(f"Error saving history: {exc}") from exc


def clear_history():
    for path in {_write_history_path(), legacy_data_path(HISTORY_FILE)}:
        if path.exists():
            path.unlink()


def export_history_entry(entry, file_path):
    results = entry.get("results", [])
    if not results:
        raise ServiceError("History entry tidak punya result untuk diexport.", title="Warning")

    first = results[0]
    if entry.get("txt_file", "").startswith("Machine:"):
        columns = ["Circuit No", "Field", "Machine Value", "Program Value", "Type", "Description"]
    elif len(first) == 6:
        columns = ["Circuit No", "Side", "Part (Reference)", "Part (Source)", "Type", "Description"]
    else:
        columns = ["Circuit No", "Part (Reference)", "Part (Source)", "Type", "Description"]

    pd.DataFrame(results, columns=columns).to_excel(file_path, index=False)

