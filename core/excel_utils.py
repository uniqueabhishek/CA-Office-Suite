"""
Shared Excel and CSV file utilities for CA Firm Office Suite.

This module provides common file I/O operations used across all tools
to eliminate code duplication and ensure consistency.
"""

import os
import re

import pandas as pd

SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv")

# Characters Excel rejects in a worksheet name, and which would be read as path
# separators in a filename or zip entry.
_UNSAFE_NAME_CHARS = re.compile(r"[\\/*?:\[\]]")


def safe_name(name):
    """
    Replace characters that are invalid in worksheet names, filenames and zip
    entries with underscores.

    Args:
        name (str): Raw name, typically a sheet name taken from a workbook

    Returns:
        str: A name safe to use in a path, zip entry or worksheet title
    """
    return _UNSAFE_NAME_CHARS.sub("_", str(name)).strip() or "Sheet"


def _engine_for(ext):
    """
    Pick the pandas Excel engine for a file extension.

    Legacy .xls needs xlrd; everything else goes through pandas' default. This
    replaces a blanket try/except that retried with xlrd on any failure, which
    hid the real error behind a second, unrelated one.

    Args:
        ext (str): Lowercased file extension, including the dot

    Returns:
        str or None: Engine name, or None to let pandas choose
    """
    return "xlrd" if ext == ".xls" else None


def list_excel_files_in_folder(folder):
    """
    Recursively lists all supported Excel and CSV files in a given folder.

    Args:
        folder (str): Root folder path to search

    Returns:
        list: List of absolute file paths for all Excel/CSV files found
    """
    out = []
    for root, _, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(SUPPORTED_EXTENSIONS):
                out.append(os.path.join(root, f))
    return out


def read_file_to_df(path):
    """
    Reads a single supported file (Excel or CSV) into a pandas DataFrame.
    For Excel files, it reads only the first sheet.

    Args:
        path (str): Path to the file to read

    Returns:
        pandas.DataFrame: DataFrame with all values as strings
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    return pd.read_excel(path, sheet_name=0, dtype=str, keep_default_na=False, engine=_engine_for(ext))


def read_all_sheets(path):
    """
    Reads all sheets from a supported file into a dictionary of DataFrames.

    Args:
        path (str): Path to the file to read

    Returns:
        dict: Dictionary mapping sheet names to DataFrames
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return {"Sheet1": pd.read_csv(path, dtype=str, keep_default_na=False)}
    return pd.read_excel(path, sheet_name=None, dtype=str, keep_default_na=False, engine=_engine_for(ext))
