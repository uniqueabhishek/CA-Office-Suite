"""
Excel formatter worker for background processing.

Handles Excel file cleaning and formatting in a separate thread to keep UI responsive.
This worker was originally FileProcessorThread in excel_formatter_tool.py.
"""

import os
import traceback

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from config.constants import MAX_SHEET_NAME_LENGTH
from core.excel_utils import read_all_sheets
from core.excel_writer import apply_formatting_to_worksheet
from workers.base_worker import BaseWorker


class FormatterWorker(BaseWorker):
    """
    Background worker for processing Excel files with cleaning and formatting.

    Applies data transformations (trim, number conversion, date normalization, etc.)
    to Excel/CSV files and saves the processed results.

    Args:
        files (list): List of file paths to process
        options (dict): Processing options (trim, numbers, dates, etc.)
        output_folder (str): Output folder path (None to overwrite originals)
        apply_transformations_func (callable): Function that applies transformations to DataFrames
        parent (QObject, optional): Parent object

    Signals:
        progress_update: Emitted with (progress_value, message)
        finished: Emitted with (success, message) when complete

    Example:
        >>> worker = FormatterWorker(files, options, output_folder, apply_all_transformations)
        >>> worker.progress_update.connect(self.on_progress)
        >>> worker.finished.connect(self.on_finished)
        >>> worker.start()
    """

    def __init__(self, files, options, output_folder, apply_transformations_func, parent=None):
        """Initialize formatter worker."""
        super().__init__(parent)
        self.files = files
        self.options = options
        self.output_folder = output_folder
        self.apply_transformations_func = apply_transformations_func

    def _output_path(self, file_path):
        """
        Return where a processed file is written.

        With no output folder the source file is overwritten in place; with one,
        the result keeps its name but a workbook always lands as .xlsx.

        Args:
            file_path (str): Path of the source file

        Returns:
            str: Destination path
        """
        if not self.output_folder:
            return file_path

        base = os.path.basename(file_path)
        name, ext = os.path.splitext(base)
        if ext.lower() == ".csv":
            return os.path.join(self.output_folder, base)
        return os.path.join(self.output_folder, base if ext.lower() == ".xlsx" else name + ".xlsx")

    def _save_processed(self, processed_sheets, out_path):
        """
        Write the processed sheets to out_path.

        Args:
            processed_sheets (dict): Sheet name -> (DataFrame, number format map)
            out_path (str): Destination path, .csv or .xlsx

        Returns:
            None
        """
        if out_path.lower().endswith(".csv"):
            if processed_sheets:
                df_proc, _ = next(iter(processed_sheets.values()))
                df_proc.to_csv(out_path, index=False)
            return

        wb = Workbook()
        # Remove default sheet
        if wb.active is not None:
            wb.remove(wb.active)

        want_formatting = (
            self.options.get("apply_autofit", False)
            or self.options.get("apply_number_format", False)
            or self.options.get("apply_theme", False)
        )

        for sheetname, (df_proc, nf_map) in processed_sheets.items():
            ws = wb.create_sheet(title=sheetname[:MAX_SHEET_NAME_LENGTH])
            # Use fast dataframe_to_rows instead of slow iterrows
            for row in dataframe_to_rows(df_proc, index=False, header=True):
                ws.append(row)

            # Format in memory before the single save (50% less I/O). Each sheet
            # is formatted with its own number format map: merging the maps, as
            # this used to, stamps one sheet's formats onto any other sheet that
            # happens to share a header name.
            if want_formatting:
                apply_formatting_to_worksheet(
                    ws,
                    number_format_map=nf_map,
                    apply_theme=self.options.get("apply_theme", False),
                    apply_autofit=self.options.get("apply_autofit", False),
                )

        # Single save operation
        wb.save(out_path)

    def run(self):
        """
        Main processing loop running in background thread.

        This method runs in a separate thread and should not be called directly.
        Use start() to begin execution.
        """
        try:
            total = len(self.files)

            for idx, file_path in enumerate(self.files, start=1):
                # Check for cancellation
                if self.is_cancelled:
                    self.finished.emit(False, "Processing cancelled by user.")
                    return

                self.emit_progress(idx, total, f"Processing: {file_path}")

                try:
                    # Read all sheets
                    sheets = read_all_sheets(file_path)
                    processed_sheets = {}

                    for sheetname, df in sheets.items():
                        # Apply all transformations using optimized pipeline
                        # Which columns were converted is not reported per sheet.
                        df_proc, _conversions, nf_map = self.apply_transformations_func(df, self.options)
                        processed_sheets[sheetname] = (df_proc, nf_map)

                    # Save processed result(s)
                    out_path = self._output_path(file_path)
                    self._save_processed(processed_sheets, out_path)

                    verb = "Saved" if self.output_folder else "Overwrote"
                    kind = "CSV" if out_path.lower().endswith(".csv") else "workbook"
                    self.progress_update.emit(idx, f"{verb} {kind}: {out_path}")

                except Exception as e:  # pylint: disable=broad-except
                    # One bad file is logged and skipped so the batch continues.
                    error_msg = f"Error processing {file_path}: {e}"
                    self.progress_update.emit(idx, error_msg)
                    self.progress_update.emit(idx, f"Traceback:\n{traceback.format_exc()}")

            # All files processed successfully
            self.finished.emit(True, f"All {total} files processed successfully.")

        except Exception as e:  # pylint: disable=broad-except
            # Last guard in a QThread body: an escaping exception would kill the
            # thread with no signal emitted, leaving the UI waiting forever.
            self.finished.emit(False, f"Unexpected error: {e}\n{traceback.format_exc()}")
