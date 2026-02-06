import os
from PyQt5 import QtWidgets
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
    QGroupBox,
)

# Import shared utilities from core modules
from core.excel_utils import list_excel_files_in_folder, read_all_sheets
from core.excel_writer import save_df_to_excel
from workers.merge_worker import MergeWorker
from ui.components import ProgressLogger

# ---------- GUI ----------
class ExcelMergeSplitWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Merge/Split Tool")
        # self.resize(800, 600)
        self.selected_files = []
        self.output_folder = None
        # Not used for Merge/Split mostly, but good for Split
        self.overwrite_originals = False
        self.worker = None  # Background worker thread

        self._build_ui()

    def _build_ui(self):
        # main = QWidget()
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)
        # self.setCentralWidget(main)

        # A. File Handling
        file_group = QGroupBox("1. File Selection")
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

        # B. Tasks
        tasks_group = QGroupBox("2. Select Operation")
        tasks_layout = QVBoxLayout()
        tasks_group.setLayout(tasks_layout)

        self.chk_merge = QCheckBox("Merge selected files into one workbook")
        merge_row = QWidget()
        mr_layout = QHBoxLayout()
        merge_row.setLayout(mr_layout)
        mr_layout.setContentsMargins(0, 0, 0, 0)
        mr_layout.addWidget(QLabel("   Merged filename:"))
        self.merge_name_edit = QtWidgets.QLineEdit("merged_output.xlsx")
        mr_layout.addWidget(self.merge_name_edit)
        tasks_layout.addWidget(self.chk_merge)
        tasks_layout.addWidget(merge_row)

        self.chk_split = QCheckBox("Split workbook sheets into separate files")
        tasks_layout.addWidget(self.chk_split)

        # C. Process
        process_group = QGroupBox("3. Process")
        process_layout = QVBoxLayout()
        process_group.setLayout(process_layout)

        btn_row = QWidget()
        btn_layout = QHBoxLayout()
        btn_row.setLayout(btn_layout)
        btn_layout.setContentsMargins(0, 0, 0, 0)

        self.process_btn = QPushButton("Start Processing")
        self.process_btn.clicked.connect(self.start_processing)
        btn_layout.addWidget(self.process_btn)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.cancel_processing)
        self.cancel_btn.setEnabled(False)
        btn_layout.addWidget(self.cancel_btn)

        process_layout.addWidget(btn_row)

        main_layout.addWidget(file_group)
        main_layout.addWidget(tasks_group)
        main_layout.addWidget(process_group)

        # Progress and log display
        self.progress_logger = ProgressLogger(log_height=200, show_label=True)
        main_layout.addWidget(self.progress_logger)

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
            self.progress_logger.log(f"Added {len(files)} files.")

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
            self.progress_logger.log(f"Added {added} files from folder {folder}.")

    def clear_files(self):
        self.selected_files = []
        self.file_list_widget.clear()
        self.progress_logger.log("Cleared file list.")

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select output folder", "")
        if folder:
            self.output_folder = folder
            self.out_folder_label.setText(f"Output folder: {folder}")
            self.progress_logger.log(f"Output folder set: {folder}")

    def start_processing(self):
        if not self.file_list_widget.count():
            QMessageBox.warning(self, "No files", "No files selected.")
            return

        if not self.chk_merge.isChecked() and not self.chk_split.isChecked():
            QMessageBox.warning(
                self, "No operation",
                "Please select Merge or Split operation.")
            return

        if not self.output_folder:
            QMessageBox.warning(
                self, "No Output Folder",
                "Please select an output folder.")
            return

        files = []
        for i in range(self.file_list_widget.count()):
            item = self.file_list_widget.item(i)
            if item is not None:
                text = item.text()
                if text:
                    files.append(text)

        # Merge Logic
        if self.chk_merge.isChecked():
            merged_name = self.merge_name_edit.text().strip() or "merged_output.xlsx"
            merged_path = os.path.join(self.output_folder, merged_name)
            self.progress_logger.log("Starting merge operation in background...")

            # Clear progress
            self.progress_logger.set_progress(0)
            self.progress_logger.set_max_progress(len(files))

            # Disable UI during processing
            self.process_btn.setEnabled(False)
            self.cancel_btn.setEnabled(True)

            # Create and start worker thread
            self.worker = MergeWorker(files, merged_path, apply_formatting=True)
            self.worker.progress_update.connect(self.on_progress_update)
            self.worker.finished.connect(self.on_merge_finished)
            self.worker.start()
            return  # Exit early - worker will handle completion

        # Split Logic
        if self.chk_split.isChecked():
            total = len(files)
            self.progress_logger.set_max_progress(total)
            self.progress_logger.set_progress(0)
            for idx, file_path in enumerate(files, start=1):
                self.progress_logger.log(f"Splitting: {file_path}")
                try:
                    sheets = read_all_sheets(file_path)
                    base_name = os.path.splitext(
                        os.path.basename(file_path))[0]
                    for sheetname, df in sheets.items():
                        out_name = f"{base_name}__{sheetname[:20]}.xlsx"
                        out_path = os.path.join(self.output_folder, out_name)
                        save_df_to_excel(df, out_path, sheet_name=sheetname)
                        self.progress_logger.log(f"  Saved sheet: {out_path}")
                except Exception as e:
                    self.progress_logger.log(f"Failed to split {file_path}: {e}")
                self.progress_logger.set_progress(idx)

        self.progress_logger.log("Processing complete.")
        QMessageBox.information(self, "Done", "Processing finished.")

    def cancel_processing(self):
        """Cancel the current merge operation."""
        if self.worker and self.worker.isRunning():
            self.progress_logger.log("Cancelling operation...")
            self.worker.cancel()

    def on_progress_update(self, progress_value, message):
        """Handle progress updates from worker thread."""
        self.progress_logger.set_progress(progress_value)
        self.progress_logger.log(message)

    def on_merge_finished(self, success, message):
        """Handle completion of merge operation."""
        self.reset_ui_after_processing()

        if success:
            QMessageBox.information(self, "Success", message)
            self.progress_logger.log(f"✓ {message}")
        else:
            QMessageBox.critical(self, "Error", message)
            self.progress_logger.log(f"✗ {message}")

    def reset_ui_after_processing(self):
        """Re-enable UI controls after processing completes."""
        self.process_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
