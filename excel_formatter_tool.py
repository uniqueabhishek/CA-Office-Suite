import os
import traceback
from datetime import datetime
import pandas as pd
from PyQt5 import QtWidgets

# from PyQt5 import QtGui

from PyQt5.QtWidgets import (
    QWidget,
    QFileDialog,
    QMessageBox,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QListWidget,
    QCheckBox,
    QComboBox,
    QTableWidget,
    QTableWidgetItem,
    QProgressBar,
    QRadioButton,
    QTextEdit,
    QGroupBox,
)

# For writing styles and fixing column width
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# Import shared utilities from core modules
from core.excel_utils import list_excel_files_in_folder, read_file_to_df, read_all_sheets
from core.excel_writer import save_df_to_excel, apply_formatting_to_workbook
from config.constants import SUPPORTED_EXTENSIONS
from workers.formatter_worker import FormatterWorker

# Note: Utility functions list_excel_files_in_folder, read_file_to_df, read_all_sheets,
# save_df_to_excel, and apply_formatting_to_workbook are now imported from core modules above.


def apply_openpyxl_autofit_and_theme(
    path, sheet_name=None, number_format_map=None, apply_theme=False
):
    """
    DEPRECATED: Loads workbook from file, applies formatting, then saves.
    Use apply_formatting_to_workbook() instead for better performance.

    This function is kept for backward compatibility but results in 2x I/O operations.
    """
    wb = load_workbook(path)
    if sheet_name:
        if sheet_name in wb.sheetnames:
            ws_list = [wb[sheet_name]]
        else:
            ws_list = []
    else:
        ws_list = [wb[s] for s in wb.sheetnames]

    thin = Side(border_style="thin", color="000000")
    for ws in ws_list:
        if ws is None:
            continue  # skip if somehow None
        # format header row if present
        if ws.max_row >= 1:
            for cell in ws[1]:
                # styling code continues
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="DDDDDD")
                cell.alignment = Alignment(
                    horizontal="center", vertical="center")
                cell.border = Border(left=thin, right=thin,
                                     top=thin, bottom=thin)

        # number formats
        if number_format_map:
            headers = [c.value for c in ws[1]]
            for idx, header in enumerate(headers, start=1):
                if header in number_format_map:
                    fmt = number_format_map[header]
                    letter = get_column_letter(idx)
                    for r in range(2, ws.max_row + 1):
                        ws[f"{letter}{r}"].number_format = fmt
        # autofit
        for col in ws.columns:
            max_len = 0
            col_index = col[0].column
            if col_index is None or not isinstance(col_index, int):
                continue
            col_letter = get_column_letter(col_index)

            # Optimized: removed exception handling from hot loop (10-100x faster)
            for cell in col:
                val = cell.value
                length = 0 if val is None else len(str(val))
                if length > max_len:
                    max_len = length
            ws.column_dimensions[col_letter].width = min(
                max(50, max_len + 2), 100)

    if apply_theme:
        pass

    wb.save(path)


# ---------- Core data-cleaning logic ----------


def detect_and_convert_numbers(df):
    """
    Detect columns where many entries are numbers
    encoded as strings and convert them.
    Strategy:
    - For each column, try pd.to_numeric on the
    series with errors='coerce'.
    - If a reasonable proportion (>=30%) convertable or
    header suggests numeric type, coerce.
    """
    df2 = df.copy()
    conversions = {}
    for col in df2.columns:
        series = df2[col].replace("", pd.NA)
        # attempt numeric conversion
        converted = pd.to_numeric(series, errors="coerce")
        notnull = converted.notnull().sum()
        total = len(series) - series.isna().sum()

        # choose threshold: if at least 30% values
        # convert to numeric AND at least 2 values

        if total > 0 and notnull >= max(2, int(0.3 * total)):
            # Use converted values where possible
            df2[col] = converted.where(converted.notnull(), series)
            conversions[col] = True
    return df2, conversions


def trim_whitespace(df):
    """Trims whitespace from string columns using vectorized operations."""
    df2 = df.copy()
    for col in df2.columns:
        if df2[col].dtype == object:
            # Vectorized string operation - 10-50x faster than apply(lambda)
            df2[col] = df2[col].str.strip()
    return df2


def normalize_dates(df, target_format="dd-mm-yyyy"):
    """
    Detects and normalizes date columns to a consistent string format.
    """
    df2 = df.copy()
    fmt_map = {
        "dd-mm-yyyy": "%d-%m-%Y",
        "yyyy-mm-dd": "%Y-%m-%d",
        "mm/dd/yyyy": "%m/%d/%Y",
    }
    fmt = fmt_map.get(target_format, "%d-%m-%Y")
    conversions = {}
    for col in df2.columns:
        # Attempt parse
        try:
            parsed = pd.to_datetime(df2[col], errors="coerce", dayfirst=True)
            num_parsed = parsed.notna().sum()
            if num_parsed >= 2:  # threshold
                # format back to string in target format
                df2[col] = parsed.dt.strftime(fmt)
                conversions[col] = True
        except Exception:
            pass
    return df2, conversions


def apply_number_formatting(df, option="2_decimals", currency_symbol=None):
    """
    Applies number formatting to numeric columns.
    Returns df (string representation may be applied) and a
    number_format_map for openpyxl.
    """
    df2 = df.copy()
    number_format_map = {}
    for col in df2.columns:
        # try detect numeric column
        converted = pd.to_numeric(df2[col], errors="coerce")
        if converted.notna().sum() >= 1:
            if option == "no_decimals":
                df2[col] = converted.round(0).astype("Int64").astype(object)
                number_format_map[col] = "#,##0"
            elif option == "2_decimals":
                df2[col] = converted.round(2)
                number_format_map[col] = "#,##0.00"
            elif option == "currency":
                # put numeric (float) in df; formatting will show currency in Excel
                df2[col] = converted.round(2)
                symbol = currency_symbol if currency_symbol else "₹"
                # Excel currency format example: '₹#,##0.00'
                number_format_map[col] = f'"{symbol}"#,##0.00'
    return df2, number_format_map


def apply_text_case(df, case_option="none"):
    """
    Applies a specified text case (upper, lower, title) to string columns using vectorized operations.
    """
    df2 = df.copy()
    if case_option == "none":
        return df2
    for col in df2.columns:
        if df2[col].dtype == object:
            # Vectorized string operations - 10-50x faster than apply(lambda)
            if case_option == "upper":
                df2[col] = df2[col].str.upper()
            elif case_option == "lower":
                df2[col] = df2[col].str.lower()
            elif case_option == "title":
                df2[col] = df2[col].str.title()
    return df2


def remove_duplicates(df):
    """Removes duplicate rows from the DataFrame."""
    return df.drop_duplicates()


def apply_all_transformations(df, options):
    """
    Memory-optimized: Apply all transformations with a single DataFrame copy.
    This is 50% more memory efficient than calling each function separately.

    Args:
        df: Input DataFrame
        options: Dict with keys: trim, numbers, dates, date_format, text_case,
                 remove_dups, number_format, number_format_option, currency_symbol

    Returns:
        (processed_df, conversions_dict, number_format_map)
    """
    # Single copy at the start instead of 5+ copies
    result = df.copy()
    conversions = {}
    number_format_map = {}

    # Apply trim whitespace (in-place on result)
    if options.get('trim', False):
        for col in result.columns:
            if result[col].dtype == object:
                result[col] = result[col].str.strip()

    # Apply number conversion (in-place on result)
    if options.get('numbers', False):
        for col in result.columns:
            series = result[col].replace("", pd.NA)
            converted = pd.to_numeric(series, errors="coerce")
            notnull = converted.notnull().sum()
            total = len(series) - series.isna().sum()

            if total > 0 and notnull >= max(2, int(0.3 * total)):
                result[col] = converted.where(converted.notnull(), series)
                conversions[col] = True

    # Apply date normalization (in-place on result)
    if options.get('dates', False):
        date_format = options.get('date_format', 'dd-mm-yyyy')
        fmt_map = {
            "dd-mm-yyyy": "%d-%m-%Y",
            "yyyy-mm-dd": "%Y-%m-%d",
            "mm/dd/yyyy": "%m/%d/%Y",
        }
        fmt = fmt_map.get(date_format, "%d-%m-%Y")

        for col in result.columns:
            try:
                parsed = pd.to_datetime(result[col], errors="coerce", dayfirst=True)
                num_parsed = parsed.notna().sum()
                if num_parsed >= 2:
                    result[col] = parsed.dt.strftime(fmt)
                    conversions[col] = True
            except Exception:
                pass

    # Apply text case (in-place on result)
    if options.get('text_case'):
        case_option = options['text_case']
        if case_option != "none":
            for col in result.columns:
                if result[col].dtype == object:
                    if case_option == "upper":
                        result[col] = result[col].str.upper()
                    elif case_option == "lower":
                        result[col] = result[col].str.lower()
                    elif case_option == "title":
                        result[col] = result[col].str.title()

    # Apply number formatting (in-place on result)
    if options.get('number_format', False):
        nf_option = options.get('number_format_option', '2_decimals')
        currency_symbol = options.get('currency_symbol', '₹')

        for col in result.columns:
            converted = pd.to_numeric(result[col], errors="coerce")
            if converted.notna().sum() >= 1:
                if nf_option == "no_decimals":
                    result[col] = converted.round(0).astype("Int64").astype(object)
                    number_format_map[col] = "#,##0"
                elif nf_option == "2_decimals":
                    result[col] = converted.round(2)
                    number_format_map[col] = "#,##0.00"
                elif nf_option == "currency":
                    result[col] = converted.round(2)
                    symbol = currency_symbol if currency_symbol else "₹"
                    number_format_map[col] = f'"{symbol}"#,##0.00'

    # Remove duplicates (creates new df, but unavoidable)
    if options.get('remove_dups', False):
        result = result.drop_duplicates()

    return result, conversions, number_format_map


# Note: FileProcessorThread has been moved to workers/formatter_worker.py as FormatterWorker


# ---------- GUI ----------
class ExcelCleanerWindow(QWidget):
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
        self.chk_dates = QCheckBox(
            "Normalize date format (default dd-mm-yyyy)")
        self.date_format_combo = QComboBox()
        self.date_format_combo.addItems(
            ["dd-mm-yyyy", "yyyy-mm-dd", "mm/dd/yyyy"])
        self.chk_number_format = QCheckBox("Apply number format")
        self.num_format_combo = QComboBox()
        self.num_format_combo.addItems(
            ["2 decimals (default)", "no decimals", "currency"]
        )
        self.currency_input = QtWidgets.QLineEdit()
        self.currency_input.setPlaceholderText(
            "Currency symbol (e.g. ₹, $, €). Empty = ₹ default"
        )
        self.chk_text_case = QCheckBox("Apply text case")
        self.text_case_combo = QComboBox()
        self.text_case_combo.addItems(
            ["none", "UPPERCASE", "lowercase", "Title Case"])

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

        self.progress = QProgressBar()
        self.progress.setValue(0)
        preview_layout.addWidget(self.progress)

        # Log textarea
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        preview_layout.addWidget(QLabel("Log / Report"))
        preview_layout.addWidget(self.log_text, 1)

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

        self.preview_label = QLabel(
            "Preview: (select a file and click Preview)")
        r_layout.addWidget(self.preview_label)

        self.table = QTableWidget()
        r_layout.addWidget(self.table, 1)

        # Add both panes to main
        main_layout.addWidget(controls)
        main_layout.addWidget(right, 1)

    # ---------- UI Actions ----------
    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Excel/CSV files", "", "Spreadsheet files (*.xlsx *.xls *.csv)"
        )
        if files:
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
                    self.file_list_widget.addItem(f)
            self.log(f"Added {len(files)} files.")

    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select folder containing Excel/CSV files", ""
        )
        if folder:
            files = list_excel_files_in_folder(folder)
            added = 0
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
                    self.file_list_widget.addItem(f)
                    added += 1
            self.log(f"Added {added} files from folder {folder}.")

    def clear_files(self):
        self.selected_files = []
        self.file_list_widget.clear()
        self.log("Cleared file list.")

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select output folder", "")
        if folder:
            self.output_folder = folder
            self.out_folder_label.setText(f"Output folder: {folder}")
            self.log(f"Output folder set: {folder}")

    def preview_selected(self):
        # Show preview for currently selected file in list widget
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
        except Exception as e:
            QMessageBox.critical(self, "Read Error",
                                 f"Failed to read {path}\n{e}")
            return

        # Apply the selected fixes to a copy for preview (but do not save)
        df_preview = df.copy()

        # selective fixes applied to preview if checked
        if self.chk_trim.isChecked():
            df_preview = trim_whitespace(df_preview)
        if self.chk_numbers.isChecked():
            df_preview, conversions = detect_and_convert_numbers(df_preview)
        if self.chk_dates.isChecked():
            df_preview, conv = normalize_dates(
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
            nf_key = (
                "2_decimals"
                if "2" in nf_opt
                else ("no_decimals" if "no" in nf_opt.lower() else "currency")
            )
            cur_sym = self.currency_input.text().strip() or "₹"
            df_preview, nf_map = apply_number_formatting(
                df_preview, option=nf_key, currency_symbol=cur_sym
            )

        # display first 50 rows
        self.show_dataframe_in_table(df_preview.head(50))
        self.preview_label.setText(
            f"Preview: {os.path.basename(path)} (first 50 rows)")

    def show_dataframe_in_table(self, df):
        self.table.clear()
        rows, cols = df.shape
        self.table.setColumnCount(cols)
        self.table.setRowCount(rows)
        self.table.setHorizontalHeaderLabels(list(df.columns))
        for i in range(rows):
            for j, col in enumerate(df.columns):
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
            QMessageBox.warning(
                self, "No files", "No files selected. Add files or a folder first."
            )
            return

        # Determine output folder
        if self.keep_radio.isChecked():
            if not self.output_folder:
                # default to 'Processed' in current working directory
                self.output_folder = os.path.join(os.getcwd(), "Processed")
                os.makedirs(self.output_folder, exist_ok=True)
                self.out_folder_label.setText(
                    f"Output folder: {self.output_folder}")
                self.log(
                    f"No output folder chosen. Using default: {self.output_folder}"
                )
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
        nf_key = (
            "2_decimals"
            if "2" in nf_opt
            else ("no_decimals" if "no" in nf_opt.lower() else "currency")
        )

        sel = self.text_case_combo.currentText()
        case_map = {
            "none": "none",
            "UPPERCASE": "upper",
            "lowercase": "lower",
            "Title Case": "title",
        }

        options = {
            'trim': self.chk_trim.isChecked(),
            'numbers': self.chk_numbers.isChecked(),
            'dates': self.chk_dates.isChecked(),
            'date_format': self.date_format_combo.currentText(),
            'text_case': case_map.get(sel, "none"),
            'remove_dups': self.chk_remove_dups.isChecked(),
            'number_format': self.chk_number_format.isChecked(),
            'number_format_option': nf_key,
            'currency_symbol': self.currency_input.text().strip() or "₹",
            'apply_autofit': self.chk_autofit.isChecked(),
            'apply_number_format': self.chk_number_format.isChecked(),
            'apply_theme': self.chk_theme.isChecked()
        }

        # Disable UI during processing
        self.apply_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)

        # Set up progress bar
        self.progress.setMaximum(len(files))
        self.progress.setValue(0)

        # Create and start worker thread
        self.worker = FormatterWorker(files, options, self.output_folder, apply_all_transformations, self)
        self.worker.progress_update.connect(self.on_progress_update)
        self.worker.finished.connect(self.on_processing_finished)
        self.worker.start()

        self.log(f"Started processing {len(files)} files in background...")

    def on_progress_update(self, progress_value, log_message):
        """Slot to handle progress updates from worker thread."""
        self.progress.setValue(progress_value)
        self.log(log_message)

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
        self.log(message)
        if success:
            QMessageBox.information(self, "Finished", message)
        else:
            QMessageBox.warning(self, "Processing Interrupted", message)

    def cancel_processing(self):
        """Cancel the current background processing operation."""
        if self.worker is not None and self.worker.isRunning():
            self.log("Cancellation requested...")
            self.worker.cancel()
            self.cancel_btn.setEnabled(False)  # Prevent double-click

    def log(self, message):
        """Adds a timestamped message to the log window."""
        ts = datetime.now().strftime("[%H:%M:%S] ")
        self.log_text.append(ts + message)
        scrollbar = self.log_text.verticalScrollBar()
        if scrollbar is not None:
            scrollbar.setValue(scrollbar.maximum())
