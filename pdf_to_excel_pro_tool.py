import os
from PyQt5.QtWidgets import (
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QScrollArea,
    QHBoxLayout,
    QFileDialog,
    QMessageBox,
    QCheckBox,
    QTextEdit,
    QFrame,
)
from PyQt5.QtCore import Qt

from ui.components import ProgressLogger
from workers.pdf_worker import PDFExtractWorker, PDFWorker


class PDFTableExtractor(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF to Excel PRO Tool")
        self.checkboxes = []
        self.tables = []
        self.pdf_path = None
        self.worker = None  # Background scan/export thread
        self.setup_ui()

    def setup_ui(self):
        main_layout = QVBoxLayout(self)

        # Top section with label and buttons
        top_layout = QHBoxLayout()

        self.label = QLabel("Select a PDF file to extract tables")
        self.label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        top_layout.addWidget(self.label)
        # Turn QLabel into a framed box
        self.label.setFrameShape(QFrame.Box)  # Box frame

        self.label.setStyleSheet("border: 1px solid grey; " "padding: 3px;")
        # Border and padding

        self.select_btn = QPushButton("Select PDF")
        self.select_btn.clicked.connect(self.select_pdf)
        top_layout.addWidget(self.select_btn)
        self.select_btn.setMaximumWidth(120)

        self.convert_btn = QPushButton("Convert to Excel")
        self.convert_btn.clicked.connect(self.convert_to_excel)
        self.convert_btn.setEnabled(False)
        top_layout.addWidget(self.convert_btn)
        self.convert_btn.setMaximumWidth(160)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.cancel_processing)
        self.cancel_btn.setEnabled(False)
        top_layout.addWidget(self.cancel_btn)
        self.cancel_btn.setMaximumWidth(100)

        main_layout.addLayout(top_layout)

        # Scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)

        self.scroll_area.setWidget(self.scroll_widget)
        main_layout.addWidget(self.scroll_area)

        # Progress bar and log
        self.progress_logger = ProgressLogger(log_height=120)
        main_layout.addWidget(self.progress_logger)

    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open PDF", "", "PDF Files (*.pdf)"
        )
        if file_path:
            self.pdf_path = file_path
            self.label.setText(f"Selected: {os.path.basename(file_path)}")
            self.convert_btn.setEnabled(False)
            self.preview_tables()

    def preview_tables(self):
        """
        Scan the selected PDF in a background thread so the UI stays responsive.
        """
        self.clear_preview()
        self.tables = []

        self.select_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.progress_logger.clear()
        self.progress_logger.set_max_progress(100)  # Scan reports percentages
        self.progress_logger.log(f"Scanning {os.path.basename(self.pdf_path)}...")

        self.worker = PDFExtractWorker(self.pdf_path, self)
        self.worker.progress_update.connect(self.on_progress_update)
        self.worker.tables_ready.connect(self.on_tables_ready)
        self.worker.finished.connect(self.on_scan_finished)
        self.worker.start()

    def on_tables_ready(self, tables):
        """Slot receiving extracted tables from the scan worker."""
        self.tables = tables
        for page_num, table_num, df in self.tables:
            checkbox = QCheckBox(f"Page {page_num} - Table {table_num}")
            checkbox.setChecked(True)
            self.checkboxes.append(checkbox)
            self.scroll_layout.addWidget(checkbox)

            preview = QTextEdit()
            preview.setReadOnly(True)
            preview.setText(df.to_string(index=False))
            self.scroll_layout.addWidget(preview)

    def on_scan_finished(self, success, message):
        """Slot handling completion of the background PDF scan."""
        self.select_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.convert_btn.setEnabled(bool(self.tables))
        self._release_worker()

        self.progress_logger.log(message)
        if not success and not self.tables:
            QMessageBox.information(self, "No Tables", message)

    def clear_preview(self):
        for i in reversed(range(self.scroll_layout.count())):
            item = self.scroll_layout.itemAt(i)
            if item is not None:
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
        self.checkboxes.clear()

    def convert_to_excel(self):
        """
        Export the checked tables to Excel in a background thread.
        """
        selected_tables = [
            tbl for tbl, chk in zip(self.tables, self.checkboxes)
            if chk.isChecked()
        ]
        if not selected_tables:
            QMessageBox.information(
                self,
                "No Selection",
                "Please select at least one table to export.",
            )
            return

        base_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
        output_path = os.path.join(
            os.path.dirname(self.pdf_path), base_name + ".xlsx"
        )

        self.select_btn.setEnabled(False)
        self.convert_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)

        self.progress_logger.set_max_progress(len(selected_tables))
        self.progress_logger.set_progress(0)

        self.worker = PDFWorker(self.pdf_path, selected_tables, output_path, self)
        self.worker.progress_update.connect(self.on_progress_update)
        self.worker.finished.connect(self.on_export_finished)
        self.worker.start()

    def on_progress_update(self, progress_value, message):
        """Slot to handle progress updates from a worker thread."""
        self.progress_logger.set_progress(progress_value)
        self.progress_logger.log(message)

    def on_export_finished(self, success, message):
        """Slot handling completion of the background Excel export."""
        self.select_btn.setEnabled(True)
        self.convert_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self._release_worker()

        self.progress_logger.log(message)
        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.critical(self, "Error", message)

    def cancel_processing(self):
        """Cancel the running background operation."""
        if self.worker is not None and self.worker.isRunning():
            self.progress_logger.log("Cancellation requested...")
            self.worker.cancel()
            self.cancel_btn.setEnabled(False)  # Prevent double-click

    def _release_worker(self):
        """Detach and schedule deletion of the finished worker."""
        if self.worker is not None:
            self.worker.deleteLater()
            self.worker = None
