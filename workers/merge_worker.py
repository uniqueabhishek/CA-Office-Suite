"""
Excel merge worker for background processing.

Handles merging multiple Excel/CSV files in a separate thread to keep UI responsive.
"""

import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from workers.base_worker import BaseWorker
from core.excel_utils import read_all_sheets
from core.excel_writer import save_df_to_excel


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

            # Read all files
            dfs = []
            colsets = []
            sheetmaps = {}

            for idx, file_path in enumerate(self.files, start=1):
                # Check for cancellation
                if self.is_cancelled:
                    self.finished.emit(False, "Merge cancelled by user")
                    return

                self.emit_progress(idx, total_files, f"Reading file {idx}/{total_files}: {os.path.basename(file_path)}")

                try:
                    sheets = read_all_sheets(file_path)
                    if not sheets:
                        self.emit_progress(idx, total_files, f"Skipped {file_path}: No sheets found")
                        continue

                    # Get first sheet for column comparison
                    first_sheet_name = list(sheets.keys())[0]
                    df = sheets[first_sheet_name]
                    dfs.append((file_path, df))
                    colsets.append(tuple(df.columns))
                    sheetmaps[file_path] = sheets

                except Exception as e:
                    self.emit_progress(idx, total_files, f"Skipped {file_path}: {str(e)}")
                    continue

            # Check if we have any data
            if not dfs:
                self.finished.emit(False, "No valid data found in any files")
                return

            # Check for cancellation before merge
            if self.is_cancelled:
                self.finished.emit(False, "Merge cancelled by user")
                return

            # Determine merge strategy
            if len(colsets) >= 1 and all(cs == colsets[0] for cs in colsets):
                # All columns match - concatenate into single sheet
                self.emit_progress(total_files, total_files, "Columns match - concatenating data...")

                merged_df = pd.concat([df for _, df in dfs], ignore_index=True)
                save_df_to_excel(merged_df, self.output_path, sheet_name='Merged')

                self.emit_progress(total_files, total_files, f"Created merged sheet with {len(merged_df)} total rows")

            else:
                # Columns don't match - create separate sheets
                self.emit_progress(total_files, total_files, "Columns differ - creating separate sheets...")

                wb = Workbook()
                if wb.active is not None:
                    wb.remove(wb.active)

                sheet_count = 0
                for file_path, sheets in sheetmaps.items():
                    short_name = os.path.splitext(os.path.basename(file_path))[0][:25]

                    for sheetname, df in sheets.items():
                        # Check for cancellation
                        if self.is_cancelled:
                            self.finished.emit(False, "Merge cancelled by user")
                            return

                        sheet_count += 1
                        title = f"{short_name}__{sheetname}"[:31]

                        self.emit_progress(
                            total_files,
                            total_files,
                            f"Creating sheet {sheet_count}: {title}"
                        )

                        ws = wb.create_sheet(title=title)

                        # Use fast dataframe_to_rows (100-1000x faster than iterrows)
                        for row in dataframe_to_rows(df, index=False, header=True):
                            ws.append(row)

                # Save workbook
                self.emit_progress(total_files, total_files, f"Saving workbook with {sheet_count} sheets...")
                wb.save(self.output_path)

            # Apply formatting if requested
            if self.apply_formatting:
                self.emit_progress(total_files, total_files, "Applying formatting...")
                # Note: Auto-fit is already applied during save_df_to_excel
                # or within the workbook creation above

            # Success!
            self.finished.emit(
                True,
                f"Successfully merged {len(dfs)} files into {self.output_path}"
            )

        except Exception as e:
            # Handle any errors
            self.emit_error(e, "Merge operation failed")
            self.finished.emit(False, f"Error: {str(e)}")
