# User Guide - CA Firm Office Suite

> **Complete guide for using all tools in the CA Office Suite**

---

## Table of Contents

1. [Getting Started](#getting-started)
2. [Excel Formatter Tool](#excel-formatter-tool)
3. [PDF to Excel Tool](#pdf-to-excel-tool)
4. [Excel Merge & Split Tool](#excel-merge-split-tool)
5. [Common Tasks](#common-tasks)
6. [Troubleshooting](#troubleshooting)
7. [Tips & Best Practices](#tips--best-practices)

---

## Getting Started

### Launching the Application

1. **Open Terminal/Command Prompt**
2. **Navigate to project folder**:
   ```bash
   cd "D:\D Desktop\PY __ CA Office Suite"
   ```
3. **Run the application**:
   ```bash
   python desktop_suite_app.py
   ```

### Main Window Overview

The application opens with a **tabbed interface** containing three tools:

| Tab | Tool Name | Purpose |
|-----|-----------|---------|
| **Tab 1** | PDF to Excel | Extract tables from PDF documents |
| **Tab 2** | Excel Formatter | Clean and format Excel/CSV files |
| **Tab 3** | Merge & Split | Combine or separate Excel files |

---

## Excel Formatter Tool

### Overview
The Excel Formatter tool automatically cleans messy Excel and CSV files by removing errors, standardizing formats, and applying professional styling.

### Step-by-Step Guide

#### Step 1: Add Files to Process

**Option A: Add Individual Files**
1. Click **"Add Files"** button
2. Select one or more `.xlsx`, `.xls`, or `.csv` files
3. Files appear in the left panel list

**Option B: Add Entire Folder**
1. Click **"Add Folder"** button
2. Select a folder containing Excel/CSV files
3. All supported files are automatically added (including subfolders)

**Option C: Clear the List**
- Click **"Clear Files"** to remove all files from the list

---

#### Step 2: Select Output Location

**Option A: Save to New Folder** (Recommended)
1. Select **"Keep original + create processed copy"** radio button
2. Click **"Select Output Folder"**
3. Choose where to save cleaned files
4. Default: Creates "Processed" folder in current directory

**Option B: Overwrite Original Files**
1. Select **"Overwrite original files"** radio button
2.   **Warning**: This permanently modifies your original files!
3. Make backups before using this option

---

#### Step 3: Choose Formatting Fixes

**A. Formatting Fixes**

| Option | What It Does | Example |
|--------|--------------|---------|
| **Convert numbers stored as text** | Detects numbers saved as text and converts them | `"12345"` ’ `12345` |
| **Trim leading/trailing spaces** | Removes extra spaces at start/end of cells | `"  Hello  "` ’ `"Hello"` |
| **Normalize date format** | Converts dates to consistent format | `12/31/2024` ’ `31-12-2024` |
| **Apply number format** | Formats numbers with decimals/commas | `1234.5` ’ `1,234.50` |
| **Apply text case** | Changes text to UPPER, lower, or Title Case | `hello world` ’ `Hello World` |

**Date Format Options**:
- `dd-mm-yyyy` (Default): 31-12-2024
- `yyyy-mm-dd`: 2024-12-31
- `mm/dd/yyyy`: 12/31/2024

**Number Format Options**:
- **2 decimals** (Default): 1,234.50
- **No decimals**: 1,235
- **Currency**: ¹1,234.50 (customize symbol in text box)

**Text Case Options**:
- **none**: No change
- **UPPERCASE**: ALL CAPS
- **lowercase**: all lowercase
- **Title Case**: First Letter Capitalized

---

**B. Advanced Tasks**

| Option | What It Does | When to Use |
|--------|--------------|-------------|
| **Remove duplicate rows** | Deletes exact duplicate rows | Clean up imported data |
| **Auto-fit column widths** | Adjusts column width to fit content | Make all text visible |
| **Apply Excel theme** | Adds professional header styling | Client-ready reports |

---

#### Step 4: Preview Changes

**Before processing all files**, preview what will happen:

1. **Select a file** from the list (click on it)
2. Click **"Preview Selected File"**
3. **Preview table** on the right shows:
   - First 50 rows
   - All selected transformations applied
   -  **No changes saved yet** - this is just a preview!

**Review the preview** to ensure:
- Dates are formatted correctly
- Numbers are recognized properly
- Text case looks good
- No data loss occurred

---

#### Step 5: Process All Files

1. Click **"Apply to All"** button
2. **Progress bar** shows processing status
3. **Log window** displays:
   - Current file being processed
   - Any errors or warnings
   - Completion status

**During Processing**:
-  **UI remains responsive** (you can see updates)
- ø Click **"Cancel"** to stop processing
- =Ê Progress bar shows completion percentage

**After Processing**:
-  Success dialog appears
- =Á Check output folder for processed files
- =Ä Log shows summary of all operations

---

### Excel Formatter - Common Use Cases

#### Use Case 1: Clean Imported CSV Data
**Scenario**: Imported bank statement CSV with messy formatting

**Steps**:
1. Add CSV file
2. Enable:
   -  Convert numbers stored as text
   -  Trim leading/trailing spaces
   -  Normalize date format
   -  Remove duplicate rows
3. Preview and apply

**Result**: Clean, properly formatted Excel file

---

#### Use Case 2: Prepare Client Report
**Scenario**: Internal Excel file needs professional formatting

**Steps**:
1. Add Excel file
2. Enable:
   -  Apply number format (2 decimals)
   -  Apply text case (Title Case)
   -  Auto-fit column widths
   -  Apply Excel theme
3. Apply to all

**Result**: Professional-looking report with consistent formatting

---

#### Use Case 3: Batch Process 100 Files
**Scenario**: Monthly processing of vendor invoices

**Steps**:
1. Add folder containing all invoice files
2. Select consistent transformations
3. Set output folder
4. Click "Apply to All"
5. **Processing time**: ~10-30 seconds for 100 files

**Result**: All files cleaned and ready for analysis

---

## PDF to Excel Tool

### Overview
Extract tables from PDF documents and save them as formatted Excel workbooks.

### Step-by-Step Guide

#### Step 1: Select PDF File
1. Click **"Select PDF"** button
2. Choose a PDF file containing tables
3. Wait for table detection (may take 10-30 seconds for large PDFs)

---

#### Step 2: Preview Detected Tables
- **Left panel** shows list of detected tables
- **Format**: `Page X - Table Y (Rows x Cols)`
- **Example**: `Page 1 - Table 1 (25 rows x 5 cols)`

**For each table**:
1. Click on table name in list
2. **Preview** appears on right showing table contents
3. Verify the table looks correct

---

#### Step 3: Select Tables to Extract

**Option A: Extract All Tables**
1. Keep all tables checked (default)
2. Click **"Convert to Excel"**

**Option B: Extract Specific Tables**
1. **Uncheck** tables you don't want
2. Only checked tables will be exported
3. Click **"Convert to Excel"**

---

#### Step 4: Convert to Excel
1. Click **"Convert to Excel"**
2. Choose output location and filename
3. Wait for conversion
4. **Success** message appears when done

**Output Format**:
- One Excel workbook (.xlsx)
- Each table becomes a separate sheet
- Sheet names: `Page1_Table1`, `Page2_Table1`, etc.
- Auto-fitted columns
- Professional header formatting

---

### PDF to Excel - Common Use Cases

#### Use Case 1: Extract Financial Statements
**Scenario**: PDF annual report with Balance Sheet and P&L tables

**Steps**:
1. Select PDF
2. Preview both tables
3. Verify numbers extracted correctly
4. Convert to Excel
5. **Manual cleanup**: Check for any merged cells or formatting issues

---

#### Use Case 2: Extract Multiple Invoices
**Scenario**: PDF with 10 pages, each page has one invoice table

**Steps**:
1. Select PDF
2. All 10 tables detected automatically
3. Preview first and last table to verify
4. Convert to Excel
5. **Result**: Single workbook with 10 sheets (one per invoice)

---

## Excel Merge & Split Tool

### Overview
Combine multiple Excel/CSV files into a single workbook, or split a workbook into separate files.

### Merge Files - Step-by-Step

#### Step 1: Add Files to Merge
1. Click **"Add Files"** or **"Add Folder"**
2. Select Excel/CSV files to merge
3. Files appear in list

---

#### Step 2: Configure Merge Options

**Merge Mode Options**:
- **Combine into single sheet**: All data stacked into one sheet
- **Keep separate sheets**: Each file becomes a separate sheet in output workbook

**Additional Options**:
- **Include headers from first file only**: Useful when merging similar files
- **Preserve original sheet names**: Keep original names instead of auto-generating

---

#### Step 3: Merge Files
1. Click **"Merge Files"**
2. Choose output location
3. Wait for merge operation
4. Check output workbook

**Result**:
- Single `.xlsx` file
- All data consolidated
- Professional formatting applied

---

### Merge Files - Common Use Cases

#### Use Case 1: Consolidate Monthly Reports
**Scenario**: 12 monthly sales reports (Jan.xlsx, Feb.xlsx, ..., Dec.xlsx)

**Merge Mode**: Keep separate sheets

**Steps**:
1. Add all 12 files
2. Select "Keep separate sheets"
3. Merge into `Annual_Sales_2024.xlsx`

**Result**: One workbook with 12 sheets (Jan, Feb, ..., Dec)

---

#### Use Case 2: Combine Regional Data
**Scenario**: Sales data from 5 regions in separate CSV files

**Merge Mode**: Combine into single sheet

**Steps**:
1. Add all regional CSV files
2. Select "Combine into single sheet"
3. Enable "Include headers from first file only"
4. Merge into `National_Sales.xlsx`

**Result**: One sheet with all regional data stacked

---

## Common Tasks

### Task 1: Clean 50 Excel Files at Once

**Time Required**: 10-30 seconds

1. **Excel Formatter** tab
2. **Add Folder** with 50 files
3. Enable common fixes:
   -  Trim whitespace
   -  Convert numbers
   -  Normalize dates
4. **Preview** one file
5. **Apply to All**
6. Check **"Processed"** folder

---

### Task 2: Extract Tables from 10-Page PDF

**Time Required**: 30-60 seconds

1. **PDF to Excel** tab
2. **Select PDF** (10 pages)
3. Wait for table detection
4. **Preview** tables to verify accuracy
5. **Uncheck** any incorrect detections
6. **Convert to Excel**
7. Open output and verify

---

### Task 3: Create Master File from 20 Workbooks

**Time Required**: 5-10 seconds

1. **Merge & Split** tab
2. **Add Files** (20 workbooks)
3. Select **"Keep separate sheets"**
4. **Merge Files**
5. Save as `Master_Workbook.xlsx`

---

## Troubleshooting

### Problem: "File failed to read"

**Possible Causes**:
- File is corrupted
- File is password-protected
- File is open in another program
- Unsupported file format

**Solutions**:
1. Close file in Excel/other programs
2. Remove password protection
3. Verify file extension (.xlsx, .xls, .csv only)
4. Try opening file in Excel to verify it's not corrupted

---

### Problem: Numbers not converting properly

**Possible Causes**:
- Numbers have special characters
- Mixed data in column (text and numbers)
- Less than 30% of values are numeric

**Solutions**:
1. **Preview** the file first
2. Check if numbers have currency symbols or units
3. Use text editor to check raw data format
4. May need manual cleanup in Excel

---

### Problem: Dates showing as numbers

**Possible Causes**:
- Dates stored as serial numbers
- Mixed date formats in same column
- Non-standard date format

**Solutions**:
1. Enable **"Normalize date format"**
2. Try different date format options
3. Preview to verify detection
4. Manual cleanup may be needed for mixed formats

---

### Problem: PDF tables not detected

**Possible Causes**:
- PDF is image-based (scanned document)
- Table has no clear borders
- Text is too small or distorted

**Solutions**:
1. Verify PDF has selectable text (not image)
2. Try different PDF (test with known-good file)
3. Use OCR software first if scanned
4. May need manual extraction

---

### Problem: Application won't start

**Possible Causes**:
- Python not installed
- Missing dependencies
- Wrong Python version

**Solutions**:
```bash
# Check Python version
python --version  # Should be 3.8+

# Reinstall dependencies
pip install -r requirements.txt

# Try running with error output
python desktop_suite_app.py
```

---

## Tips & Best Practices

### Performance Tips

 **DO**:
- Close other applications when processing large files
- Use "Add Folder" for batch operations
- Preview before processing all files
- Process files in batches of 50-100

L **DON'T**:
- Open files in Excel while processing
- Process files on network drives (slow)
- Run multiple instances simultaneously

---

### Data Safety Tips

 **DO**:
- Always create backups before using "Overwrite"
- Test transformations on sample file first
- Preview changes before applying to all files
- Keep original files in separate folder

L **DON'T**:
- Use "Overwrite" without backups
- Process files directly in production folders
- Skip the preview step
- Assume all data is perfect after processing

---

### Efficiency Tips

**Reuse Settings**:
- Once you find settings that work, document them
- Use same settings for similar files
- Create standard operating procedures (SOPs)

**Keyboard Shortcuts**:
- `Tab`: Navigate between controls
- `Space`: Toggle checkboxes
- `Enter`: Activate focused button

**File Organization**:
```
Project/
   Original_Files/          # Keep originals here
   Processed/               # Auto-generated output
   Archive/                 # Move completed files here
```

---

### Quality Control Tips

**Always verify**:
1.  Check first file output manually
2.  Spot-check random files in batch
3.  Verify row counts match (no data loss)
4.  Check formulas still work (if any)
5.  Validate totals and summaries

**Common verification methods**:
- Compare file sizes (similar ± formatting)
- Check row count: Original vs Processed
- Verify key totals in summary rows
- Open in Excel and scan for errors

---

## Appendix

### Supported File Formats

| Format | Extension | Read | Write | Notes |
|--------|-----------|------|-------|-------|
| Excel 2007+ | .xlsx |  |  | Recommended |
| Excel 97-2003 | .xls |  | L | Converted to .xlsx |
| CSV | .csv |  |  | UTF-8 encoding |
| PDF | .pdf |  | L | Table extraction only |

### File Size Limits

| Operation | Recommended | Maximum | Notes |
|-----------|-------------|---------|-------|
| Excel read | < 50 MB | 100 MB | Performance degrades |
| Excel write | < 50 MB | 100 MB | May take 30+ seconds |
| PDF read | < 20 MB | 50 MB | Depends on page count |
| Batch processing | 50 files | 200 files | Process in batches |

### System Requirements

**Minimum**:
- 4 GB RAM
- 2 GB free disk space
- Python 3.8+

**Recommended for large files**:
- 8 GB RAM
- 5 GB free disk space
- SSD for faster I/O

---

## Getting Help

**Documentation**:
- [README.md](README.md) - Project overview
- [ARCHITECTURE.md](ARCHITECTURE.md) - Technical details
- [DEVELOPER_GUIDE.md](DEVELOPER_GUIDE.md) - For contributors

**Support**:
- GitHub Issues: Report bugs or request features
- Email: support@ca-office-suite.com
- Community: Discord server (link in README)

---

**User Guide Version**: 2.0
**Last Updated**: December 2024
**For**: CA Firm Office Suite v2.0
