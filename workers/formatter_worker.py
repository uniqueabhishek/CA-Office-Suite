"""
Excel formatter worker for background processing.

Handles Excel file cleaning and formatting in a separate thread to keep UI responsive.
This worker was originally FileProcessorThread in excel_formatter_tool.py.
"""

import os
import traceback

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from core.excel_utils import read_all_sheets
from core.excel_writer import apply_formatting_to_workbook
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
                        df_proc, conversions, nf_map = self.apply_transformations_func(df, self.options)
                        processed_sheets[sheetname] = (df_proc, nf_map)

                    # Save processed result(s)
                    if self.output_folder:
                        # Create same base filename in output folder
                        base = os.path.basename(file_path)
                        name, ext = os.path.splitext(base)

                        if ext.lower() == ".csv":
                            # For CSV: write single sheet
                            if processed_sheets:
                                df_proc, nf_map = list(processed_sheets.values())[0]
                                out_path = os.path.join(self.output_folder, base)
                                df_proc.to_csv(out_path, index=False)
                                self.progress_update.emit(idx, f"Saved CSV: {out_path}")
                        else:
                            # Create workbook with sheets
                            out_path = os.path.join(
                                self.output_folder, base if base.lower().endswith(".xlsx") else name + ".xlsx"
                            )
                            wb = Workbook()
                            # Remove default sheet
                            if wb.active is not None:
                                wb.remove(wb.active)

                            for sheetname, (df_proc, _) in processed_sheets.items():
                                ws = wb.create_sheet(title=sheetname[:31])
                                # Use fast dataframe_to_rows instead of slow iterrows
                                for row in dataframe_to_rows(df_proc, index=False, header=True):
                                    ws.append(row)

                            # Apply formatting in memory BEFORE saving (50% less I/O)
                            if (
                                self.options.get("apply_autofit", False)
                                or self.options.get("apply_number_format", False)
                                or self.options.get("apply_theme", False)
                            ):
                                # Build merged number_format_map across sheets
                                merged_nf = {}
                                for _, nfmap in processed_sheets.values():
                                    if nfmap:
                                        merged_nf.update(nfmap)
                                # Apply formatting to workbook in memory
                                apply_formatting_to_workbook(
                                    wb,
                                    number_format_map=merged_nf,
                                    apply_theme=self.options.get("apply_theme", False),
                                    apply_autofit=self.options.get("apply_autofit", False),
                                )

                            # Single save operation
                            wb.save(out_path)
                            self.progress_update.emit(idx, f"Saved workbook: {out_path}")
                    else:
                        # Overwrite original file
                        if file_path.lower().endswith(".csv"):
                            if processed_sheets:
                                df_proc, nf_map = list(processed_sheets.values())[0]
                                df_proc.to_csv(file_path, index=False)
                                self.progress_update.emit(idx, f"Overwrote CSV: {file_path}")
                        else:
                            wb = Workbook()
                            if wb.active is not None:
                                wb.remove(wb.active)

                            for sheetname, (df_proc, _) in processed_sheets.items():
                                ws = wb.create_sheet(title=sheetname[:31])
                                # Use fast dataframe_to_rows instead of slow iterrows
                                for row in dataframe_to_rows(df_proc, index=False, header=True):
                                    ws.append(row)

                            # Apply formatting in memory BEFORE saving
                            if (
                                self.options.get("apply_autofit", False)
                                or self.options.get("apply_number_format", False)
                                or self.options.get("apply_theme", False)
                            ):
                                merged_nf = {}
                                for _, nfmap in processed_sheets.values():
                                    if nfmap:
                                        merged_nf.update(nfmap)
                                apply_formatting_to_workbook(
                                    wb,
                                    number_format_map=merged_nf,
                                    apply_theme=self.options.get("apply_theme", False),
                                    apply_autofit=self.options.get("apply_autofit", False),
                                )

                            # Single save operation
                            wb.save(file_path)
                            self.progress_update.emit(idx, f"Overwrote workbook: {file_path}")

                except Exception as e:
                    error_msg = f"Error processing {file_path}: {e}"
                    self.progress_update.emit(idx, error_msg)
                    self.progress_update.emit(idx, f"Traceback:\n{traceback.format_exc()}")

            # All files processed successfully
            self.finished.emit(True, f"All {total} files processed successfully.")

        except Exception as e:
            # Unexpected error in thread
            self.finished.emit(False, f"Unexpected error: {e}\n{traceback.format_exc()}")
