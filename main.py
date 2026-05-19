import sys
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from PySide6.QtWidgets import QApplication

from themes.theme_manager import ThemeManager
from ui.main_window import MainWindow
from utils.paths import set_windows_app_id


def main():
    set_windows_app_id()
    app = QApplication(sys.argv)
    app.setOrganizationName("SMTTools")
    app.setApplicationName("BomComparatorQt")

    theme_manager = ThemeManager(app)
    window = MainWindow(theme_manager)
    theme_manager.apply()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

