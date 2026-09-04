"""
File merging logic shared by the desktop suite and the Flask web app.

Extracted from excel_merge_split_tool.py so both front-ends produce
identical output from the same code path.
"""

import os
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from config.constants import MAX_SHEET_NAME_LENGTH
from core.excel_utils import read_all_sheets
from core.excel_writer import save_df_to_excel, apply_formatting_to_workbook


def merge_files_logic(files, merged_path):
    """
    Merge multiple files into one workbook.

    If every file shares the same column set the rows are concatenated into a
    single 'Merged' sheet; otherwise each source sheet is written to its own
    sheet in the output workbook.

    Args:
        files (list): Paths of the files to merge
        merged_path (str): Path the merged workbook is written to

    Returns:
        None
    """
    dfs = []
    colsets = []
    sheetmaps = {}

    # Reading Phase
    for f in files:
        try:
            sheets = read_all_sheets(f)
            if not sheets:
                continue
            # prefer first sheet for column comparison
            first_sheet_name = list(sheets.keys())[0]
            df = sheets[first_sheet_name]
            dfs.append((f, df))
            colsets.append(tuple(df.columns))
            sheetmaps[f] = sheets
        except Exception as e:  # pylint: disable=broad-except
            # A single unreadable file should not abort the whole merge
            print(f"Skipping {f} during merge: {e}")

    # Decision Phase: Concatenate or Separate Sheets?
    # if all column sets identical, concat
    if len(colsets) >= 1 and all(cs == colsets[0] for cs in colsets):
        merged_df = pd.concat([df for _, df in dfs], ignore_index=True)
        # Using save_df_to_excel from core, which handles writer logic
        save_df_to_excel(merged_df, merged_path, sheet_name='Merged')
    else:
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

    # Re-open to apply the same styling the desktop app applies
    wb = load_workbook(merged_path)
    apply_formatting_to_workbook(wb, apply_autofit=True)
    wb.save(merged_path)
