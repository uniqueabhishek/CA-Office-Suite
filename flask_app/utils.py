import pdfplumber
import pandas as pd
import io
import os

def extract_tables_from_pdf(pdf_path):
    """
    Extracts tables from a PDF file.
    Returns a list of dictionaries: [{'page': 1, 'id': 1, 'data': dataframe}, ...]
    """
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            page_tables = page.extract_tables()
            for j, table in enumerate(page_tables):
                if table:
                    # Clean up table data
                    # Handle cases where table might not have headers or messy data
                    # For consistency with original script: use first row as header
                    if len(table) > 1:
                        headers = table[0]
                        # Deduplicate headers
                        seen = {}
                        new_headers = []
                        for col in headers:
                            if col is None:
                                col = ""
                            col = str(col).strip()
                            if col in seen:
                                seen[col] += 1
                                new_headers.append(f"{col}.{seen[col]}")
                            else:
                                seen[col] = 0
                                new_headers.append(col)

                        df = pd.DataFrame(table[1:], columns=new_headers)
                    else:
                        df = pd.DataFrame(table) # Fallback if only 1 row

                    tables.append({
                        'page': i + 1,
                        'id': j + 1,
                        'html': df.to_html(classes='table table-striped table-bordered table-hover', index=False),
                        'data': df.to_json(orient='split') # Using split safe for duplicates if any remain, but mostly to match structure
                    })
    return tables

def convert_selected_tables_to_excel(pdf_path, selected_indices):
    """
    selected_indices: list of strings "page-table_id" e.g. ["1-1", "3-2"]
    """
    output = io.BytesIO()

    # We need to re-extract to get the actual dataframes to write
    # Optimization: In a real prod app, we might cache the dataframes.
    # For this scale, re-extracting specifically or filtering is fine.

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        with pdfplumber.open(pdf_path) as pdf:
            tables_found = 0
            for i, page in enumerate(pdf.pages):
                page_tables = page.extract_tables()
                for j, table in enumerate(page_tables):
                    table_id_str = f"{i+1}-{j+1}"

                    if table_id_str in selected_indices:
                        if table and len(table) > 0:
                            if len(table) > 1:
                                headers = table[0]
                                # Deduplicate headers
                                seen = {}
                                new_headers = []
                                for col in headers:
                                    if col is None:
                                        col = ""
                                    col = str(col).strip()
                                    if col in seen:
                                        seen[col] += 1
                                        new_headers.append(f"{col}.{seen[col]}")
                                    else:
                                        seen[col] = 0
                                        new_headers.append(col)

                                df = pd.DataFrame(table[1:], columns=new_headers)
                            else:
                                df = pd.DataFrame(table)

                            sheet_name = f"Page{i+1}_Table{j+1}"
                            # Excel sheet names len limit is 31
                            if len(sheet_name) > 31:
                                sheet_name = sheet_name[:31]

                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            tables_found += 1

            # If no tables were found/selected (shouldn't happen if logic is correct), write a dummy
            if tables_found == 0:
                pd.DataFrame({'Info': ['No tables selected']}).to_excel(writer, sheet_name='Info')

    output.seek(0)
    return output
