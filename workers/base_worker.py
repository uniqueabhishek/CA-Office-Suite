"""
Base worker class for background processing threads.

Provides common functionality for all worker threads including:
- Progress reporting
- Error handling
- Cancellation support
"""

import traceback

from PyQt5.QtCore import QThread, pyqtSignal


class BaseWorker(QThread):
    """
    Base class for all background processing workers.

    Provides standard signals and cancellation support that all workers inherit.

    Signals:
        progress_update (int, str): Emits (progress_value, log_message)
        finished (bool, str): Emits (success, final_message)

    Attributes:
        is_cancelled (bool): Flag indicating if cancellation was requested
    """

    # Signals for communication with main thread
    progress_update = pyqtSignal(int, str)  # (progress_value, log_message)
    finished = pyqtSignal(bool, str)  # (success, final_message)

    def __init__(self, parent=None):
        """
        Initialize the base worker.

        Args:
            parent: Parent QObject (usually None for workers)
        """
        super().__init__(parent)
        self.is_cancelled = False

    def cancel(self):
        """
        Request cancellation of the current operation.

        Sets the is_cancelled flag. The run() method should check this flag
        periodically and exit gracefully when True.
        """
        self.is_cancelled = True

    def emit_progress(self, current, total, message):  # pylint: disable=unused-argument
        """
        Helper method to emit progress updates.

        The scale is not transmitted: each tool calls
        ProgressLogger.set_max_progress() itself before starting a worker, so
        the bar already knows the total and only needs the current value.
        'total' is kept because it makes the call sites state the scale the
        emitted value belongs to, which is what has to match that maximum.

        Args:
            current (int): Current progress value, on the scale of 'total'
            total (int): Total items to process, for the reader's benefit
            message (str): Progress message to display

        Example:
            >>> self.emit_progress(5, 10, "Processing file 5 of 10")
        """
        self.progress_update.emit(current, message)

    def emit_error(self, error, context=""):
        """
        Helper method to emit error messages with traceback.

        Args:
            error (Exception): The exception that occurred
            context (str): Optional context about where the error occurred

        Example:
            >>> try:
            >>>     risky_operation()
            >>> except Exception as e:
            >>>     self.emit_error(e, "Failed to process file")
        """
        error_msg = f"{context}: {error}" if context else str(error)
        self.progress_update.emit(0, error_msg)
        self.progress_update.emit(0, f"Traceback:\n{traceback.format_exc()}")

    def run(self):
        """
        Main worker thread execution method.

        Subclasses MUST override this method to implement their specific processing logic.

        The implementation should:
        1. Check self.is_cancelled periodically
        2. Emit progress_update signals for status updates
        3. Emit finished signal when complete or on error
        4. Handle exceptions and use emit_error() for reporting

        Example:
            >>> def run(self):
            >>>     try:
            >>>         for i, item in enumerate(self.items):
            >>>             if self.is_cancelled:
            >>>                 self.finished.emit(False, "Cancelled by user")
            >>>                 return
            >>>             self.emit_progress(i+1, len(self.items), f"Processing {item}")
            >>>             process_item(item)
            >>>         self.finished.emit(True, "All items processed")
            >>>     except Exception as e:
            >>>         self.emit_error(e, "Processing failed")
            >>>         self.finished.emit(False, str(e))
        """
        raise NotImplementedError("Subclasses must implement run() method")
