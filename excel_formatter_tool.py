"""
Excel Formatter tab of the desktop suite.

Builds the Qt form for the cleaning options and hands batches to
FormatterWorker. The transformations themselves live in core/transformations.py
so this tool and the web formatter apply identical rules.
"""

import os

import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

# Import shared utilities from core modules
from config.constants import DEFAULT_CURRENCY_SYMBOL, DEFAULT_OUTPUT_FOLDER, MAX_PREVIEW_ROWS
from core.excel_utils import list_excel_files_in_folder, read_file_to_df
from core.transformations import (
    apply_all_transformations,
    apply_number_formatting,
    apply_text_case,
    detect_and_convert_numbers,
    normalize_dates,
    trim_whitespace,
)
from ui.components import ProgressLogger
from workers.formatter_worker import FormatterWorker

# from PyQt5 import QtGui


# ---------- Core data-cleaning logic ----------
# The transformations live in core/transformations.py so the desktop app and
# the Flask web app share one implementation. Imported above.


# ---------- GUI ----------
class ExcelCleanerWindow(QWidget):
    """
    Excel Formatter window: file list, cleaning options, preview and batch run.

    A Qt form keeps one attribute per control, so the instance-attribute count
    tracks how many widgets the layout has rather than any real complexity.
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Formatter")
        # self.resize(1100, 700)
        self.selected_files = []
        self.output_folder = None
        self.overwrite_originals = False
        self.worker = None  # Background processing thread

        self._build_ui()

    def _build_ui(self):
        """Construct and lay out every widget in the window."""
        # main = QWidget()
        main_layout = QHBoxLayout()
        self.setLayout(main_layout)
        # self.setCentralWidget(main)

        # Left: Controls
        controls = QWidget()
        controls.setMinimumWidth(360)
        c_layout = QVBoxLayout()
        controls.setLayout(c_layout)

        # File selection group
        file_group = QGroupBox("A. File Handling")
        fg_layout = QVBoxLayout()
        file_group.setLayout(fg_layout)

        self.file_list_widget = QListWidget()
        fg_layout.addWidget(self.file_list_widget)

        btn_row = QWidget()
        br_layout = QHBoxLayout()
        btn_row.setLayout(br_layout)
        br_layout.setContentsMargins(0, 0, 0, 0)
        self.add_files_btn = QPushButton("Add Files")
        self.add_folder_btn = QPushButton("Add Folder")
        self.clear_files_btn = QPushButton("Clear")
        br_layout.addWidget(self.add_files_btn)
        br_layout.addWidget(self.add_folder_btn)
        br_layout.addWidget(self.clear_files_btn)
        fg_layout.addWidget(btn_row)

        self.add_files_btn.clicked.connect(self.add_files)
        self.add_folder_btn.clicked.connect(self.add_folder)
        self.clear_files_btn.clicked.connect(self.clear_files)

        # Output folder
        out_row = QWidget()
        or_layout = QHBoxLayout()
        out_row.setLayout(or_layout)
        or_layout.setContentsMargins(0, 0, 0, 0)
        self.out_folder_label = QLabel("Output folder: (not set)")
        self.select_out_btn = QPushButton("Select Output Folder")
        or_layout.addWidget(self.out_folder_label)
        or_layout.addWidget(self.select_out_btn)
        fg_layout.addWidget(out_row)
        self.select_out_btn.clicked.connect(self.select_output_folder)

        # Overwrite radio
        ow_row = QWidget()
        ow_layout = QHBoxLayout()
        ow_row.setLayout(ow_layout)
        ow_layout.setContentsMargins(0, 0, 0, 0)
        self.overwrite_radio = QRadioButton("Overwrite originals")
        self.keep_radio = QRadioButton("Save to output folder (recommended)")
        self.keep_radio.setChecked(True)
        ow_layout.addWidget(self.keep_radio)
        ow_layout.addWidget(self.overwrite_radio)
        fg_layout.addWidget(ow_row)

        # B. Formatting fixes
        fix_group = QGroupBox("B. Formatting Fixes")
        fix_layout = QVBoxLayout()
        fix_group.setLayout(fix_layout)

        self.chk_numbers = QCheckBox("Convert numbers stored as text")
        self.chk_trim = QCheckBox("Trim leading/trailing spaces")
        self.chk_dates = QCheckBox("Normalize date format (default dd-mm-yyyy)")
        self.date_format_combo = QComboBox()
        self.date_format_combo.addItems(["dd-mm-yyyy", "yyyy-mm-dd", "mm/dd/yyyy"])
        self.chk_number_format = QCheckBox("Apply number format")
        self.num_format_combo = QComboBox()
        self.num_format_combo.addItems(["2 decimals (default)", "no decimals", "currency"])
        self.currency_input = QtWidgets.QLineEdit()
        self.currency_input.setPlaceholderText("Currency symbol (e.g. ₹, $, €). Empty = ₹ default")
        self.chk_text_case = QCheckBox("Apply text case")
        self.text_case_combo = QComboBox()
        self.text_case_combo.addItems(["none", "UPPERCASE", "lowercase", "Title Case"])

        fix_layout.addWidget(self.chk_numbers)
        fix_layout.addWidget(self.chk_trim)
        date_row = QWidget()
        dr_layout = QHBoxLayout()
        date_row.setLayout(dr_layout)
        dr_layout.setContentsMargins(0, 0, 0, 0)
        dr_layout.addWidget(self.chk_dates)
        dr_layout.addWidget(self.date_format_combo)
        fix_layout.addWidget(date_row)

        nf_row = QWidget()
        nf_layout = QHBoxLayout()
        nf_row.setLayout(nf_layout)
        nf_layout.setContentsMargins(0, 0, 0, 0)
        nf_layout.addWidget(self.chk_number_format)
        nf_layout.addWidget(self.num_format_combo)
        fix_layout.addWidget(nf_row)
        fix_layout.addWidget(self.currency_input)

        tc_row = QWidget()
        tc_layout = QHBoxLayout()
        tc_row.setLayout(tc_layout)
        tc_layout.setContentsMargins(0, 0, 0, 0)
        tc_layout.addWidget(self.chk_text_case)
        tc_layout.addWidget(self.text_case_combo)
        fix_layout.addWidget(tc_row)

        # C. Basic tasks
        tasks_group = QGroupBox("C. Basic Tasks")
        tasks_layout = QVBoxLayout()
        tasks_group.setLayout(tasks_layout)

        self.chk_remove_dups = QCheckBox("Remove duplicate rows")
        tasks_layout.addWidget(self.chk_remove_dups)

        self.chk_autofit = QCheckBox("Auto-adjust column widths")
        tasks_layout.addWidget(self.chk_autofit)

        self.chk_theme = QCheckBox("Apply simple style/theme")
        tasks_layout.addWidget(self.chk_theme)

        # D. Preview & Apply
        preview_group = QGroupBox("D. Preview & Apply")
        preview_layout = QVBoxLayout()
        preview_group.setLayout(preview_layout)

        btns_preview = QWidget()
        bp_layout = QHBoxLayout()
        btns_preview.setLayout(bp_layout)
        bp_layout.setContentsMargins(0, 0, 0, 0)
        self.preview_btn = QPushButton("Preview Selected File")
        self.apply_btn = QPushButton("Apply to All")
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setEnabled(False)  # Disabled until processing starts
        bp_layout.addWidget(self.preview_btn)
        bp_layout.addWidget(self.apply_btn)
        bp_layout.addWidget(self.cancel_btn)
        preview_layout.addWidget(btns_preview)

        # Progress and log display
        self.progress_logger = ProgressLogger(log_height=200, show_label=True)
        preview_layout.addWidget(self.progress_logger, 1)

        # Wire preview and apply
        self.preview_btn.clicked.connect(self.preview_selected)
        self.apply_btn.clicked.connect(self.apply_to_all)
        self.cancel_btn.clicked.connect(self.cancel_processing)

        # Add groups to control layout
        c_layout.addWidget(file_group)
        c_layout.addWidget(fix_group)
        c_layout.addWidget(tasks_group)
        c_layout.addWidget(preview_group)
        c_layout.addStretch(1)

        # Right: Preview table
        right = QWidget()
        r_layout = QVBoxLayout()
        right.setLayout(r_layout)

        self.preview_label = QLabel("Preview: (select a file and click Preview)")
        r_layout.addWidget(self.preview_label)

        self.table = QTableWidget()
        r_layout.addWidget(self.table, 1)

        # Add both panes to main
        main_layout.addWidget(controls)
        main_layout.addWidget(right, 1)

    # ---------- UI Actions ----------
    def add_files(self):
        """Add files chosen from a dialog to the selection list."""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Excel/CSV files", "", "Spreadsheet files (*.xlsx *.xls *.csv)"
        )
        if files:
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
                    self.file_list_widget.addItem(f)
            self.progress_logger.log(f"Added {len(files)} files.")

    def add_folder(self):
        """Add every supported file found under a chosen folder."""
        folder = QFileDialog.getExistingDirectory(self, "Select folder containing Excel/CSV files", "")
        if folder:
            files = list_excel_files_in_folder(folder)
            added = 0
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
                    self.file_list_widget.addItem(f)
                    added += 1
            self.progress_logger.log(f"Added {added} files from folder {folder}.")

    def clear_files(self):
        """Empty the selection list."""
        self.selected_files = []
        self.file_list_widget.clear()
        self.progress_logger.log("Cleared file list.")

    def select_output_folder(self):
        """Choose where processed files are written instead of overwriting."""
        folder = QFileDialog.getExistingDirectory(self, "Select output folder", "")
        if folder:
            self.output_folder = folder
            self.out_folder_label.setText(f"Output folder: {folder}")
            self.progress_logger.log(f"Output folder set: {folder}")

    def preview_selected(self):
        """Show the highlighted file with the checked fixes applied, unsaved."""
        item = self.file_list_widget.currentItem()
        if not item:
            QMessageBox.warning(
                self,
                "No file selected",
                "Please select one file from the list to preview.",
            )
            return
        path = item.text()
        try:
            df = read_file_to_df(path)
        except Exception as e:  # pylint: disable=broad-except
            # Any unreadable file becomes a dialog rather than a crashed preview.
            QMessageBox.critical(self, "Read Error", f"Failed to read {path}\n{e}")
            return

        # Apply the selected fixes to a copy for preview (but do not save)
        df_preview = df.copy()

        # selective fixes applied to preview if checked
        if self.chk_trim.isChecked():
            df_preview = trim_whitespace(df_preview)
        if self.chk_numbers.isChecked():
            # The preview only shows the values; which columns were converted
            # is reported on the batch run, not here.
            df_preview, _conversions = detect_and_convert_numbers(df_preview)
        if self.chk_dates.isChecked():
            df_preview, _date_conversions = normalize_dates(
                df_preview, target_format=self.date_format_combo.currentText()
            )
        if self.chk_text_case.isChecked():
            sel = self.text_case_combo.currentText()
            case_map = {
                "none": "none",
                "UPPERCASE": "upper",
                "lowercase": "lower",
                "Title Case": "title",
            }
            df_preview = apply_text_case(df_preview, case_map.get(sel, "none"))

        # number formatting preview won't change actual numbers in preview (we'll show numbers), but we could round
        if self.chk_number_format.isChecked():
            nf_opt = self.num_format_combo.currentText()
            nf_key = "2_decimals" if "2" in nf_opt else ("no_decimals" if "no" in nf_opt.lower() else "currency")
            cur_sym = self.currency_input.text().strip() or DEFAULT_CURRENCY_SYMBOL
            # The number format map targets openpyxl cells, which the Qt table
            # preview does not have; only the rounded values are shown.
            df_preview, _nf_map = apply_number_formatting(df_preview, option=nf_key, currency_symbol=cur_sym)

        self.show_dataframe_in_table(df_preview.head(MAX_PREVIEW_ROWS))
        self.preview_label.setText(f"Preview: {os.path.basename(path)} (first {MAX_PREVIEW_ROWS} rows)")

    def show_dataframe_in_table(self, df):
        """Render a DataFrame into the preview table widget."""
        self.table.clear()
        rows, cols = df.shape
        self.table.setColumnCount(cols)
        self.table.setRowCount(rows)
        self.table.setHorizontalHeaderLabels(list(df.columns))
        for i in range(rows):
            for j in range(cols):
                val = df.iloc[i, j]
                if pd.isna(val):
                    text = ""
                else:
                    text = str(val)
                item = QTableWidgetItem(text)
                self.table.setItem(i, j, item)
        self.table.resizeColumnsToContents()

    def apply_to_all(self):
        """
        Optimized: Uses background thread for processing to keep UI responsive.
        Enables cancel button during processing.
        """
        if not self.file_list_widget.count():
            QMessageBox.warning(self, "No files", "No files selected. Add files or a folder first.")
            return

        # Determine output folder
        if self.keep_radio.isChecked():
            if not self.output_folder:
                # default to 'Processed' in current working directory
                self.output_folder = os.path.join(os.getcwd(), DEFAULT_OUTPUT_FOLDER)
                os.makedirs(self.output_folder, exist_ok=True)
                self.out_folder_label.setText(f"Output folder: {self.output_folder}")
                self.progress_logger.log(f"No output folder chosen. Using default: {self.output_folder}")
        else:
            # overwrite originals
            self.output_folder = None

        # Build a list of file paths
        files = []
        for i in range(self.file_list_widget.count()):
            item = self.file_list_widget.item(i)
            if item is not None:
                text = item.text()
                if text:
                    files.append(text)

        if not files:
            QMessageBox.warning(self, "No files", "No valid files to process.")
            return

        # Build options dictionary from UI controls
        nf_opt = self.num_format_combo.currentText()
        nf_key = "2_decimals" if "2" in nf_opt else ("no_decimals" if "no" in nf_opt.lower() else "currency")

        sel = self.text_case_combo.currentText()
        case_map = {
            "none": "none",
            "UPPERCASE": "upper",
            "lowercase": "lower",
            "Title Case": "title",
        }

        options = {
            "trim": self.chk_trim.isChecked(),
            "numbers": self.chk_numbers.isChecked(),
            "dates": self.chk_dates.isChecked(),
            "date_format": self.date_format_combo.currentText(),
            "text_case": case_map.get(sel, "none"),
            "remove_dups": self.chk_remove_dups.isChecked(),
            "number_format": self.chk_number_format.isChecked(),
            "number_format_option": nf_key,
            "currency_symbol": self.currency_input.text().strip() or DEFAULT_CURRENCY_SYMBOL,
            "apply_autofit": self.chk_autofit.isChecked(),
            "apply_number_format": self.chk_number_format.isChecked(),
            "apply_theme": self.chk_theme.isChecked(),
        }

        # Disable UI during processing
        self.apply_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)

        # Set up progress bar
        self.progress_logger.set_max_progress(len(files))
        self.progress_logger.set_progress(0)

        # Create and start worker thread
        self.worker = FormatterWorker(files, options, self.output_folder, apply_all_transformations, self)
        self.worker.progress_update.connect(self.on_progress_update)
        self.worker.finished.connect(self.on_processing_finished)
        self.worker.start()

        self.progress_logger.log(f"Started processing {len(files)} files in background...")

    def on_progress_update(self, progress_value, log_message):
        """Slot to handle progress updates from worker thread."""
        self.progress_logger.set_progress(progress_value)
        self.progress_logger.log(log_message)

    def on_processing_finished(self, success, message):
        """Slot to handle completion from worker thread."""
        # Re-enable UI
        self.apply_btn.setEnabled(True)
        self.preview_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)

        # Clean up worker
        if self.worker is not None:
            self.worker.deleteLater()
            self.worker = None

        # Show completion message
        self.progress_logger.log(message)
        if success:
            QMessageBox.information(self, "Finished", message)
        else:
            QMessageBox.warning(self, "Processing Interrupted", message)

    def cancel_processing(self):
        """Cancel the current background processing operation."""
        if self.worker is not None and self.worker.isRunning():
            self.progress_logger.log("Cancellation requested...")
            self.worker.cancel()
            self.cancel_btn.setEnabled(False)  # Prevent double-click
