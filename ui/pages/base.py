from PySide6.QtGui import QKeySequence
from PySide6.QtWidgets import QMessageBox, QProgressBar, QWidget

from workers.task_runner import TaskWorker


class WorkerPage(QWidget):
    def __init__(self, thread_pool, theme_manager, parent=None):
        super().__init__(parent)
        self.thread_pool = thread_pool
        self.theme_manager = theme_manager
        self._workers = []
        self._tables = []
        self._busy_widgets = []

    def register_model(self, model):
        self._tables.append(model)
        model.set_theme(self.theme_manager.theme)

    def register_busy_widgets(self, *widgets):
        self._busy_widgets.extend(widgets)

    def set_busy(self, busy, text=None):
        for widget in self._busy_widgets:
            widget.setDisabled(busy)
        progress = getattr(self, "progress", None)
        status_label = getattr(self, "status_label", None)
        if isinstance(progress, QProgressBar):
            progress.setVisible(busy)
            if busy:
                progress.setRange(0, 0)
            else:
                progress.setRange(0, 100)
                progress.setValue(0)
        if status_label is not None and text is not None:
            status_label.setText(text)

    def run_worker(self, fn, on_result, busy_text="Processing..."):
        worker = TaskWorker(fn)
        worker._busy_text = busy_text
        self._workers.append(worker)
        worker.signals.started.connect(lambda: self.set_busy(True, busy_text))
        worker.signals.result.connect(on_result)
        worker.signals.error.connect(self._show_worker_error)
        worker.signals.finished.connect(lambda w=worker: self._finish_worker(w))
        self.thread_pool.start(worker)

    def _finish_worker(self, worker):
        if worker in self._workers:
            self._workers.remove(worker)
        self.set_busy(False)
        status_label = getattr(self, "status_label", None)
        if status_label is not None and status_label.text() == getattr(worker, "_busy_text", ""):
            status_label.setText("")

    def _show_worker_error(self, title, detail):
        message = detail.split("\n\nTraceback", 1)[0]
        QMessageBox.warning(self, title, message)

    def apply_theme_to_models(self, theme):
        for model in self._tables:
            model.set_theme(theme)

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            focused = self.focusWidget()
            if focused and hasattr(focused, "copy"):
                focused.copy()
                return
        super().keyPressEvent(event)
