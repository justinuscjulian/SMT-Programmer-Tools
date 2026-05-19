import os
import sys
from pathlib import Path


IS_FROZEN = getattr(sys, "frozen", False)
APP_ROOT = Path(__file__).resolve().parents[1]
RUNTIME_ROOT = Path(sys.executable).resolve().parent if IS_FROZEN else APP_ROOT
PROJECT_ROOT = APP_ROOT.parent
LEGACY_ROOT = PROJECT_ROOT / "BomAppCompare_Remake"


def resource_path(relative_path):
    base_path = Path(getattr(sys, "_MEIPASS", APP_ROOT))
    return str(base_path / relative_path)


def app_data_path(filename):
    return RUNTIME_ROOT / filename


def legacy_data_path(filename):
    return LEGACY_ROOT / filename


def set_windows_app_id(app_id="JustinusCJ.BOMComparator.Qt.1.0"):
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass
