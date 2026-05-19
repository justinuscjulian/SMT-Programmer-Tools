from PySide6.QtWidgets import QLabel


class StatusBadge(QLabel):
    def __init__(self, text="0", status="INFO", parent=None):
        super().__init__(text, parent)
        self.setObjectName("StatusBadge")
        self.setMinimumHeight(28)
        self.set_status(status)

    def set_status(self, status):
        self.setProperty("status", status)
        self.style().unpolish(self)
        self.style().polish(self)

    def set_value(self, text, status=None):
        self.setText(text)
        if status:
            self.set_status(status)
