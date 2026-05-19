# SMT Programmer Tools

PySide6 remake of the original CustomTkinter app in `BomAppCompare_Remake`.

## Run From Source

```powershell
cd "C:\Users\User\Documents\PROJECT\SMT Tools"
python -m pip install -r BomAppCompare_Qt\requirements.txt
python BomAppCompare_Qt\main.py
```

## Structure

- `main.py` - application entrypoint.
- `ui/` - main window, sidebar, top bar, and pages.
- `widgets/` - reusable UI widgets.
- `themes/` - reusable dark/light theme tokens and QSS.
- `services/` - existing data-processing logic moved out of UI.
- `workers/` - `QRunnable` task execution.
- `models/` - `QAbstractTableModel` table models.
- `utils/` - sorting, path, and encoding helpers.
- `assets/` - icon/logo assets.

## Notes

- Existing CustomTkinter source is untouched.
- `history.json` format is preserved. If the new app has no local history yet, it reads the legacy history from `BomAppCompare_Remake/history.json`.
- Heavy file load, compare, export, and sync actions are dispatched through Qt workers so the UI stays responsive.

