"""
File merging logic shared by the desktop suite and the Flask web app.

Extracted from excel_merge_split_tool.py so both front-ends produce
identical output from the same code path.
"""

import logging
import os

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from config.constants import MAX_SHEET_NAME_LENGTH
from core.excel_utils import read_all_sheets
from core.excel_writer import apply_formatting_to_workbook, save_df_to_excel

logger = logging.getLogger(__name__)


def merge_files_logic(files, merged_path, on_progress=None, should_cancel=None, apply_formatting=True):
    """
    Merge multiple files into one workbook.

    If every file shares the same column set the rows are concatenated into a
    single 'Merged' sheet; otherwise each source sheet is written to its own
    sheet in the output workbook.

    The optional hooks let a background worker report progress and stop early
    without keeping its own copy of this logic.

    Args:
        files (list): Paths of the files to merge
        merged_path (str): Path the merged workbook is written to
        on_progress (callable): Optional callback(done, total, message)
        should_cancel (callable): Optional predicate polled between files; when
            it returns True the merge stops before anything is written
        apply_formatting (bool): Whether to run the final styling pass over the
            merged workbook. A concatenated sheet is styled by save_df_to_excel
            either way; this controls the pass across every sheet.

    Returns:
        bool: True when a workbook was written, False when cancelled

    Raises:
        ValueError: None of the files could be read
    """

    def report(done, message):
        if on_progress is not None:
            on_progress(done, len(files), message)

    def cancelled():
        return should_cancel is not None and should_cancel()

    dfs = []
    colsets = []
    sheetmaps = {}

    # Reading Phase
    for idx, f in enumerate(files, start=1):
        if cancelled():
            return False
        report(idx, f"Reading file {idx}/{len(files)}: {os.path.basename(f)}")
        try:
            sheets = read_all_sheets(f)
            if not sheets:
                report(idx, f"Skipped {os.path.basename(f)}: no sheets found")
                continue
            # prefer first sheet for column comparison
            first_sheet_name = list(sheets.keys())[0]
            df = sheets[first_sheet_name]
            dfs.append((f, df))
            colsets.append(tuple(df.columns))
            sheetmaps[f] = sheets
        except Exception:  # pylint: disable=broad-except
            # A single unreadable file should not abort the whole merge
            logger.warning("Skipping %s during merge: unreadable", f, exc_info=True)
            report(idx, f"Skipped {os.path.basename(f)}: unreadable")

    if not dfs:
        # Every input failed to read. Falling through would build a workbook
        # with no sheets, which openpyxl refuses to save with an error that
        # says nothing about the real cause.
        raise ValueError("None of the selected files could be read, so there is nothing to merge.")

    if cancelled():
        return False

    # Decision Phase: Concatenate or Separate Sheets?
    # if all column sets identical, concat
    if all(cs == colsets[0] for cs in colsets):
        report(len(files), "Columns match - concatenating data...")
        merged_df = pd.concat([df for _, df in dfs], ignore_index=True)
        # Using save_df_to_excel from core, which handles writer logic
        save_df_to_excel(merged_df, merged_path, sheet_name="Merged")
    else:
        report(len(files), "Columns differ - creating separate sheets...")
        # create workbook with each file as sheet
        wb = Workbook()
        if wb.active is not None:
            wb.remove(wb.active)
        for f, sheets in sheetmaps.items():
            short = os.path.splitext(os.path.basename(f))[0][:25]
            for sheetname, df in sheets.items():
                title = f"{short}__{sheetname}"[:MAX_SHEET_NAME_LENGTH]
                ws = wb.create_sheet(title=title)
                # Use fast dataframe_to_rows instead of slow iterrows
                for row in dataframe_to_rows(df, index=False, header=True):
                    ws.append(row)
        wb.save(merged_path)

    if apply_formatting:
        # Re-open to apply the same styling the desktop app applies
        report(len(files), "Applying formatting...")
        wb = load_workbook(merged_path)
        apply_formatting_to_workbook(wb, apply_autofit=True)
        wb.save(merged_path)

    return True
