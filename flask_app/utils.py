import pdfplumber
import pandas as pd
import io
import os
import zipfile
from flask import current_app

# Import core modules
from core.excel_utils import read_all_sheets, read_file_to_df
from core.excel_writer import save_df_to_excel, apply_formatting_to_workbook

# Import extracted parity logic
from core.transformations import (
    detect_and_convert_numbers,
    trim_whitespace,
    normalize_dates,
    apply_text_case,
    apply_all_transformations
)
from core.merge_logic import merge_files_logic

# --- BRIDGE FUNCTIONS ---

def process_excel_file(filepath, original_filename, options):
    """
    Process a single excel file based on options.
    Returns (BytesIO stream, output_filename)
    """
    sheets = read_all_sheets(filepath)

    # In-memory workbook
    output = io.BytesIO()

    # We will use ExcelWriter with openpyxl
    # Note: apply_all_transformations returns (df, conversions, nf_map)
    # The desktop app processes sheet by sheet.

    processed_sheets = {}

    for sheet_name, df in sheets.items():
        # Use simple or optimized pipeline?
        # The extracted code has 'apply_all_transformations'. Let's use that for exact match.
        df_proc, conversions, nf_map = apply_all_transformations(df, options)
        processed_sheets[sheet_name] = (df_proc, nf_map)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, (df_proc, _) in processed_sheets.items():
             df_proc.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    # Rewind to read for openpyxl formatting
    output.seek(0)

    # Calculate merged number format map
    merged_nf = {}
    for _, (_, nf_map) in processed_sheets.items():
        if nf_map:
            merged_nf.update(nf_map)

    # Apply Formatting using core writer logic
    from openpyxl import load_workbook
    wb = load_workbook(output)

    apply_formatting_to_workbook(
        wb,
        number_format_map=merged_nf,
        apply_theme=options.get('apply_theme', False),
        apply_autofit=options.get('apply_autofit', False)
    )

    # Save final
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    base, _ = os.path.splitext(original_filename)
    return final_output, f"{base}_processed.xlsx"


def merge_files(file_paths):
    """
    Merges multiple files into one workbook using Desktop logic.
    """
    # merge_files_logic writes to a file path usually.
    # We need to adapt it to write to BytesIO or temp file.
    # Since our merge_logic.py takes a path, we'll use a temp file.

    import uuid
    import shutil

    temp_output_path = os.path.join(os.path.dirname(file_paths[0]), f"merged_temp_{uuid.uuid4()}.xlsx")

    try:
        # Call the shared logic directly
        merge_files_logic(file_paths, temp_output_path)

        # Read back into memory
        with open(temp_output_path, 'rb') as f:
            data = f.read()

        return io.BytesIO(data)
    finally:
        # Cleanup
        if os.path.exists(temp_output_path):
            os.remove(temp_output_path)


def split_files_to_zip(file_paths, original_names):
    """
    Splits the FIRST file in the list (web update limit) into sheets and zips them.
    (Desktop logic uses read_all_sheets + save_df_to_excel loop, which is what we do here too)
    """
    import zipfile

    # Only process first file for now to match strict constraint of single response
    fpath = file_paths[0]
    fname = original_names[0]

    sheets = read_all_sheets(fpath)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        base, _ = os.path.splitext(fname)

        for sname, df in sheets.items():
            # Save sheet to bytes
            sheet_buffer = io.BytesIO()
            # Use core saver? It saves to disk.
            # We need pure bytes. dataframe to excel is standard pandas.
            df.to_excel(sheet_buffer, index=False)
            sheet_buffer.seek(0)

            zip_file.writestr(f"{base}_{sname}.xlsx", sheet_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer


# --- PDF LOGIC (EXISTING) ---
def extract_tables_from_pdf(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            page_tables = page.extract_tables()
            for j, table in enumerate(page_tables):
                if table:
                    headers = table[0]
                    seen = {}
                    new_headers = []
                    for col in headers:
                        col = str(col).strip() if col else ""
                        if col in seen:
                            seen[col] += 1
                            new_headers.append(f"{col}.{seen[col]}")
                        else:
                            seen[col] = 0
                            new_headers.append(col)
                    df = pd.DataFrame(table[1:], columns=new_headers)

                    tables.append({
                        'page': i + 1,
                        'id': j + 1,
                        'html': df.to_html(classes='table table-striped', index=False),
                        'data': df.to_json(orient='split')
                    })
    return tables

def convert_selected_tables_to_excel(pdf_path, selected_indices):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                page_tables = page.extract_tables()
                for j, table in enumerate(page_tables):
                    if f"{i+1}-{j+1}" in selected_indices:
                        if table:
                            df = pd.DataFrame(table[1:], columns=table[0]) if len(table)>1 else pd.DataFrame(table)
                            df.to_excel(writer, sheet_name=f"Page{i+1}_Table{j+1}", index=False)
    output.seek(0)
    return output
