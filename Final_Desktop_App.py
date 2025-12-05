import sys
import os
import pdfplumber
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QScrollArea,
    QHBoxLayout,
    QFileDialog,
    QMessageBox,
    QDesktopWidget,
    QCheckBox,
    QTextEdit,
    QFrame,
)
from PyQt5.QtCore import Qt


class PDFTableExtractor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Professional PDF to Excel Converter")
        self.resize(1600, 800)
        self.checkboxes = []
        self.center_window()
        self.setup_ui()

    def center_window(self):
        screen_geometry = QDesktopWidget().availableGeometry()
        screen_center = screen_geometry.center()
        frame_geometry = self.frameGeometry()
        frame_geometry.moveCenter(screen_center)
        self.move(frame_geometry.topLeft())

    def setup_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        main_layout = QVBoxLayout(main_widget)

        # Top section with label and buttons
        top_layout = QHBoxLayout()

        self.label = QLabel("Select a PDF file to extract tables")
        self.label.setAlignment(Qt.AlignLeft)
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

        main_layout.addLayout(top_layout)

        # Scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)

        self.scroll_area.setWidget(self.scroll_widget)
        main_layout.addWidget(self.scroll_area)

    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open PDF", "", "PDF Files (*.pdf)"
        )
        if file_path:
            self.pdf_path = file_path
            self.label.setText(f"Selected: {os.path.basename(file_path)}")
            self.convert_btn.setEnabled(True)
            self.preview_tables()

    def preview_tables(self):
        self.clear_preview()
        self.tables = self.extract_tables(self.pdf_path)
        if not self.tables:
            QMessageBox.information(self, "No Tables",
                                    "No tables found in this PDF.")
            return

        for idx, (page_num, table_num, df) in enumerate(self.tables):
            checkbox = QCheckBox(f"Page {page_num} - Table {table_num}")
            checkbox.setChecked(True)
            self.checkboxes.append(checkbox)
            self.scroll_layout.addWidget(checkbox)

            preview = QTextEdit()
            preview.setReadOnly(True)
            preview.setText(df.to_string(index=False))
            self.scroll_layout.addWidget(preview)

    def clear_preview(self):
        for i in reversed(range(self.scroll_layout.count())):
            widget = self.scroll_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()
        self.checkboxes.clear()

    def convert_to_excel(self):
        try:
            selected_tables = [
                tbl for tbl, chk in zip(self.tables, self.checkboxes)
                if chk.isChecked()
            ]
            if not selected_tables:
                QMessageBox.information(
                    self,
                    "No Selection",
                    "Please select at least one " "table to export.",
                )
                return

            base_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
            output_path = os.path.join(
                os.path.dirname(self.pdf_path), base_name + ".xlsx"
            )

            self.save_tables_to_excel(selected_tables, output_path)
            QMessageBox.information(
                self, "Success", f"Selected tables saved to {output_path}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def extract_tables(self, pdf_path):
        tables = []
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                page_tables = page.extract_tables()
                for j, table in enumerate(page_tables):
                    if table:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        tables.append((i + 1, j + 1, df))
        return tables

    def save_tables_to_excel(self, tables, output_path):
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for page_num, table_num, df in tables:
                sheet_name = f"Page{page_num}_Table{table_num}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFTableExtractor()
    window.show()
    sys.exit(app.exec_())
