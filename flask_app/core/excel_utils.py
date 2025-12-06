"""
Shared Excel and CSV file utilities for CA Firm Office Suite.

This module provides common file I/O operations used across all tools
to eliminate code duplication and ensure consistency.
"""

import os
import pandas as pd


SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv")


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
    # read smartly
    if ext == ".csv":
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        # For Excel, read all sheets when needed; for preview read first sheet
        # Use engine autodetection
        try:
            return pd.read_excel(path, sheet_name=0, dtype=str, keep_default_na=False)
        except Exception:
            # fallback: try with xlrd engine for old .xls
            return pd.read_excel(
                path, sheet_name=0, engine="xlrd", dtype=str, keep_default_na=False
            )


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
    else:
        return pd.read_excel(path, sheet_name=None, dtype=str, keep_default_na=False)
