import traceback

from PySide6.QtCore import QObject, QRunnable, Signal, Slot


class WorkerSignals(QObject):
    started = Signal()
    progress = Signal(int, str)
    result = Signal(object)
    error = Signal(str, str)
    finished = Signal()


class TaskWorker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    @Slot()
    def run(self):
        self.signals.started.emit()
        try:
            result = self.fn(*self.args, **self.kwargs)
            self.signals.result.emit(result)
        except Exception as exc:
            title = getattr(exc, "title", exc.__class__.__name__)
            message = getattr(exc, "message", str(exc))
            detail = traceback.format_exc()
            self.signals.error.emit(title, f"{message}\n\n{detail}")
        finally:
            self.signals.finished.emit()

