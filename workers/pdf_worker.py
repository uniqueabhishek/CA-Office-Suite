"""
PDF extraction worker for background processing.

Handles PDF table extraction in a separate thread to keep UI responsive.
"""

import pandas as pd
from workers.base_worker import BaseWorker


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
                    self.emit_progress(
                        idx,
                        total_tables,
                        f"Extracting table {idx}/{total_tables}: {sheet_name}"
                    )

                    # Write table to Excel
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Final progress update
                self.emit_progress(
                    total_tables,
                    total_tables,
                    f"Saved {total_tables} tables to {self.output_path}"
                )

            # Success!
            self.finished.emit(True, f"Successfully extracted {total_tables} tables to {self.output_path}")

        except Exception as e:
            # Handle any errors
            self.emit_error(e, "PDF extraction failed")
            self.finished.emit(False, f"Error: {str(e)}")
