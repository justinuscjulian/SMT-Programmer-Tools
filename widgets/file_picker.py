from PySide6.QtCore import Signal
from PySide6.QtWidgets import QHBoxLayout, QLabel, QLineEdit, QPushButton, QWidget


class FilePicker(QWidget):
    browse_requested = Signal()

    def __init__(self, label, button_text="Browse", parent=None):
        super().__init__(parent)
        self._path = ""
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        self.label = QLabel(label)
        self.label.setMinimumWidth(150)
        self.input = QLineEdit()
        self.input.setReadOnly(True)
        self.button = QPushButton(button_text)
        self.button.clicked.connect(self.browse_requested.emit)

        layout.addWidget(self.label)
        layout.addWidget(self.input, 1)
        layout.addWidget(self.button)

    def set_path(self, path, display_text=None):
        self._path = path or ""
        self.input.setText(display_text or self._path)

    def path(self):
        return self._path

    def clear(self):
        self.set_path("")

