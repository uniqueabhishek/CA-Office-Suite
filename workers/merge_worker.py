"""
Excel merge worker for background processing.

Handles merging multiple Excel/CSV files in a separate thread to keep UI responsive.
"""

from core.merge_logic import merge_files_logic
from workers.base_worker import BaseWorker


class MergeWorker(BaseWorker):
    """
    Background worker for merging Excel/CSV files.

    Merges multiple files into a single Excel workbook, either by concatenating
    data (if columns match) or creating separate sheets for each file.

    Args:
        files (list): List of file paths to merge
        output_path (str): Path where merged Excel file should be saved
        apply_formatting (bool): Whether to apply auto-fit and formatting
        parent (QObject, optional): Parent object

    Signals:
        progress_update: Emitted with (progress_value, message)
        finished: Emitted with (success, message) when complete

    Example:
        >>> worker = MergeWorker(file_list, output_path)
        >>> worker.progress_update.connect(self.on_progress)
        >>> worker.finished.connect(self.on_finished)
        >>> worker.start()
    """

    def __init__(self, files, output_path, apply_formatting=True, parent=None):
        """Initialize merge worker."""
        super().__init__(parent)
        self.files = files
        self.output_path = output_path
        self.apply_formatting = apply_formatting

    def run(self):
        """
        Merge files in background thread.

        This method runs in a separate thread and should not be called directly.
        Use start() to begin execution.
        """
        try:
            total_files = len(self.files)

            if total_files == 0:
                self.finished.emit(False, "No files to merge")
                return

            self.emit_progress(0, total_files, f"Starting merge of {total_files} files...")

            # The merge itself lives in core.merge_logic so the desktop and web
            # front-ends produce identical output. This worker only supplies the
            # progress and cancellation hooks the UI needs.
            try:
                completed = merge_files_logic(
                    self.files,
                    self.output_path,
                    on_progress=self.emit_progress,
                    should_cancel=lambda: self.is_cancelled,
                    apply_formatting=self.apply_formatting,
                )
            except ValueError as exc:
                self.finished.emit(False, str(exc))
                return

            if not completed:
                self.finished.emit(False, "Merge cancelled by user")
                return

            # Success!
            self.finished.emit(True, f"Successfully merged {total_files} files into {self.output_path}")

        except Exception as e:  # noqa: BLE001
            # Last guard in a QThread body: an escaping exception would kill the
            # thread with no signal emitted, leaving the UI waiting forever.
            self.emit_error(e, "Merge operation failed")
            self.finished.emit(False, f"Error: {e!s}")
