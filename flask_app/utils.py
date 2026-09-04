"""
Bridge between the Flask routes and the shared core/ processing modules.

Routes deal in uploaded files and in-memory responses; these helpers adapt the
path-based core logic to the streams Flask needs to send back.
"""

import io
import os
import uuid
import zipfile

import pandas as pd
import pdfplumber
from openpyxl import load_workbook

from config.constants import MAX_SHEET_NAME_LENGTH
from core.excel_utils import read_all_sheets, safe_name
from core.excel_writer import apply_formatting_to_workbook
from core.merge_logic import merge_files_logic
from core.transformations import apply_all_transformations

# --- BRIDGE FUNCTIONS ---


def process_excel_file(filepath, original_filename, options):
    """
    Process a single excel file based on options.
    Returns (BytesIO stream, output_filename)
    """
    sheets = read_all_sheets(filepath)

    # In-memory workbook
    output = io.BytesIO()

    processed_sheets = {}

    for sheet_name, df in sheets.items():
        # Same transformation pipeline the desktop formatter uses
        df_proc, _conversions, nf_map = apply_all_transformations(df, options)
        processed_sheets[sheet_name] = (df_proc, nf_map)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, (df_proc, _) in processed_sheets.items():
            df_proc.to_excel(writer, sheet_name=sheet_name[:MAX_SHEET_NAME_LENGTH], index=False)

    # Rewind to read for openpyxl formatting
    output.seek(0)

    # Calculate merged number format map
    merged_nf = {}
    for _, (_, nf_map) in processed_sheets.items():
        if nf_map:
            merged_nf.update(nf_map)

    # Apply Formatting using core writer logic
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
    Merges multiple files into one workbook using the shared desktop logic.

    merge_files_logic writes to a path, so the result is staged in a temp file
    and read back into memory for the response.
    """
    temp_output_path = os.path.join(
        os.path.dirname(file_paths[0]), f"merged_temp_{uuid.uuid4()}.xlsx"
    )

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
    Split every uploaded workbook into one .xlsx per sheet and return a zip.

    A response can only carry a single file, so each sheet becomes an entry in
    one archive. Entry names are sanitised and de-duplicated, since sheet names
    may contain characters that are illegal in a zip path or may collide across
    two uploads that share a filename.
    """
    zip_buffer = io.BytesIO()
    used_names = set()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for fpath, fname in zip(file_paths, original_names):
            sheets = read_all_sheets(fpath)
            base = safe_name(os.path.splitext(fname)[0])

            for sname, df in sheets.items():
                # Save sheet to bytes
                sheet_buffer = io.BytesIO()
                df.to_excel(sheet_buffer, index=False)
                sheet_buffer.seek(0)

                entry = f"{base}_{safe_name(sname)}.xlsx"
                suffix = 1
                while entry in used_names:
                    suffix += 1
                    entry = f"{base}_{safe_name(sname)}_{suffix}.xlsx"
                used_names.add(entry)

                zip_file.writestr(entry, sheet_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer


# --- PDF LOGIC ---


def extract_tables_from_pdf(pdf_path):
    """Extract every table in the PDF as preview HTML plus JSON data."""
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
    """Write the tables identified by 'page-table' ids into one workbook."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                page_tables = page.extract_tables()
                for j, table in enumerate(page_tables):
                    if f"{i+1}-{j+1}" in selected_indices and table:
                        df = (
                            pd.DataFrame(table[1:], columns=table[0])
                            if len(table) > 1
                            else pd.DataFrame(table)
                        )
                        df.to_excel(writer, sheet_name=f"Page{i+1}_Table{j+1}", index=False)
    output.seek(0)
    return output
