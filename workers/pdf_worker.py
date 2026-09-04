"""
PDF extraction workers for background processing.

Handles PDF table detection and export in separate threads to keep the UI
responsive. PDFExtractWorker scans a PDF for tables; PDFWorker writes the
selected tables to an Excel workbook.
"""

import pandas as pd
import pdfplumber
from PyQt5.QtCore import pyqtSignal

from core.excel_utils import dedupe_headers
from workers.base_worker import BaseWorker


class PDFExtractWorker(BaseWorker):
    """
    Background worker that scans a PDF and extracts every table it finds.

    pdfplumber's page parsing is slow enough to freeze the UI on large PDFs,
    so it runs here instead of on the GUI thread.

    Args:
        pdf_path (str): Path to the PDF file to scan
        parent (QObject, optional): Parent object

    Signals:
        tables_ready (object): Emitted with the list of (page_num, table_num, df)
        progress_update: Emitted with (progress_value, message)
        finished: Emitted with (success, message) when complete

    Example:
        >>> worker = PDFExtractWorker(pdf_path)
        >>> worker.tables_ready.connect(self.on_tables_ready)
        >>> worker.start()
    """

    tables_ready = pyqtSignal(object)  # list of (page_num, table_num, DataFrame)

    def __init__(self, pdf_path, parent=None):
        """Initialize PDF scanning worker."""
        super().__init__(parent)
        self.pdf_path = pdf_path

    def run(self):
        """
        Scan the PDF for tables in a background thread.

        This method runs in a separate thread and should not be called directly.
        Use start() to begin execution.
        """
        try:
            tables = []
            with pdfplumber.open(self.pdf_path) as pdf:
                total_pages = len(pdf.pages)
                if total_pages == 0:
                    self.tables_ready.emit([])
                    self.finished.emit(False, "This PDF has no pages.")
                    return

                # Page count is only known once the PDF is open, so progress is
                # reported as a percentage against a fixed 0-100 scale.
                self.emit_progress(0, 100, f"Scanning {total_pages} pages...")

                for i, page in enumerate(pdf.pages):
                    if self.is_cancelled:
                        self.finished.emit(False, "PDF scan cancelled by user")
                        return

                    for j, table in enumerate(page.extract_tables()):
                        if table:
                            df = pd.DataFrame(table[1:], columns=dedupe_headers(table[0]))
                            tables.append((i + 1, j + 1, df))

                    self.emit_progress(
                        int((i + 1) / total_pages * 100),
                        100,
                        f"Scanned page {i + 1}/{total_pages} - {len(tables)} tables so far",
                    )

            self.tables_ready.emit(tables)

            if not tables:
                self.finished.emit(False, "No tables found in this PDF.")
            else:
                self.finished.emit(True, f"Found {len(tables)} tables.")

        except Exception as e:  # noqa: BLE001
            self.emit_error(e, "PDF scan failed")
            self.tables_ready.emit([])
            self.finished.emit(False, f"Error: {e!s}")


class PDFWorker(BaseWorker):
    """
    Background worker for extracting tables from PDF files.

    Extracts selected tables from a PDF and saves them to an Excel workbook,
    all in a background thread to prevent UI freezing.

    Args:
        pdf_path (str): Path to the PDF file
        selected_tables (list): List of (page_num, table_num, df) tuples to export
        output_path (str): Path where the Excel file should be saved
        parent (QObject, optional): Parent object

    Signals:
        progress_update: Emitted with (progress_value, message)
        finished: Emitted with (success, message) when complete

    Example:
        >>> worker = PDFWorker(pdf_path, selected_tables, output_path)
        >>> worker.progress_update.connect(self.on_progress)
        >>> worker.finished.connect(self.on_finished)
        >>> worker.start()
    """

    def __init__(self, pdf_path, selected_tables, output_path, parent=None):
        """Initialize PDF extraction worker."""
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.selected_tables = selected_tables
        self.output_path = output_path

    def run(self):
        """
        Extract tables from PDF in background thread.

        This method runs in a separate thread and should not be called directly.
        Use start() to begin execution.
        """
        try:
            total_tables = len(self.selected_tables)

            if total_tables == 0:
                self.finished.emit(False, "No tables selected for export")
                return

            self.emit_progress(0, total_tables, f"Starting extraction of {total_tables} tables...")

            # Check for cancellation
            if self.is_cancelled:
                self.finished.emit(False, "PDF extraction cancelled by user")
                return

            # Create Excel writer
            self.emit_progress(0, total_tables, f"Creating Excel file: {self.output_path}")

            with pd.ExcelWriter(self.output_path, engine="openpyxl") as writer:
                for idx, (page_num, table_num, df) in enumerate(self.selected_tables, start=1):
                    # Check for cancellation before each table
                    if self.is_cancelled:
                        self.finished.emit(False, "PDF extraction cancelled by user")
                        return

                    # Create sheet name
                    sheet_name = f"Page{page_num}_Table{table_num}"

                    # Emit progress
                    self.emit_progress(idx, total_tables, f"Extracting table {idx}/{total_tables}: {sheet_name}")

                    # Write table to Excel
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Final progress update
                self.emit_progress(total_tables, total_tables, f"Saved {total_tables} tables to {self.output_path}")

            # Success!
            self.finished.emit(True, f"Successfully extracted {total_tables} tables to {self.output_path}")

        except Exception as e:  # noqa: BLE001
            # Last guard in a QThread body: an escaping exception would kill the
            # thread with no signal emitted, leaving the UI waiting forever.
            self.emit_error(e, "PDF extraction failed")
            self.finished.emit(False, f"Error: {e!s}")
