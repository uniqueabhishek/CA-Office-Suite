# Changelog

All notable changes to the CA Firm Office Suite project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.1.0] - 2026-09-04

### Fixed
- **Column widths in every generated workbook.** The autofit clamp in
  `core/excel_writer.py` read `min(max(50, max_len + 2), 100)`, which forced a
  minimum width of 50 characters on every column of every file the formatter,
  merge and split tools produced. Widths now track content and are clamped to
  `MIN_COLUMN_WIDTH`..`MAX_COLUMN_WIDTH` (10..60).
- **PDF extraction froze the desktop UI.** `pdf_to_excel_pro_tool.py` ran
  pdfplumber on the GUI thread. It now uses `PDFExtractWorker` to scan and
  `PDFWorker` to export, both in the background, with a progress bar, log and
  Cancel button. `workers/pdf_worker.py` had existed unused since Phase 4.2.
- **Web split silently discarded files.** `split_files_to_zip()` processed only
  the first upload. It now splits every uploaded workbook, sanitises zip entry
  names, and de-duplicates colliding names.
- **Unsanitised sheet names.** Sheet names containing `/ \ * ? : [ ]` produced
  malformed zip entries and could break the desktop split's output paths. Both
  paths now go through `core.excel_utils.safe_name()`.
- `PDFTableExtractor` no longer leaves `pdf_path` and `tables` undefined until
  the first scan.

### Security
- **Removed the hardcoded Flask secret key.** `SECRET_KEY` is read from the
  environment; startup fails when `FLASK_ENV=production` and it is unset, and a
  development fallback prints a warning.
- **Uploads are cleaned up.** Temporary uploads are deleted once a request is
  served, and a retention sweep removes files abandoned mid-flow. Previously
  the upload folder grew without bound.
- **Upload validation.** `/formatter` and `/merge` now reject anything that is
  not `.xlsx`, `.xls` or `.csv` before writing to disk; the merged filename is
  passed through `secure_filename()`.

### Changed
- **Single shared `core/` package.** `flask_app/core/excel_utils.py` and
  `excel_writer.py` were byte-identical copies of the root `core/` versions,
  and ~210 lines of transformation logic were duplicated between
  `excel_formatter_tool.py` and `flask_app/core/transformations.py`. The web
  app's copies were removed; `transformations.py` and `merge_logic.py` moved to
  `core/`, and both front-ends now import the same modules.
- `config/constants.py` is now actually used, and carries the column-width,
  sheet-name and upload settings.
- The upload folder is anchored to `flask_app/`, not the working directory.
- `Tally API/` renamed to `tally_api/` so it can be imported: a directory name
  containing a space cannot be a Python package, its relative import had no
  package to resolve against, and the smoke script referenced a module
  (`tally_integration`) that does not exist. `requests` was added to the
  dependencies it needs.
- Removed the duplicate `flask_app/Procfile` and `flask_app/requirements.txt`;
  the root `Procfile` and `pyproject.toml` are authoritative.

### Added
- **Test suite** (`tests/`, 89 tests) covering the core modules, the Flask
  bridge helpers and the web routes, with regression tests for the column-width
  and split-drops-files bugs. `pyproject.toml` configures pytest.
- `core.excel_utils.safe_name()` for sheet, file and zip-entry names.

### Documentation
- Repaired mojibake in six markdown files: emoji had been written as invalid
  byte pairs, `->` arrows and rupee signs as invalid bytes, and the ASCII tree
  and sequence diagrams had been overwritten with NUL padding. All docs are now
  valid UTF-8 and the diagrams have been redrawn.
- Corrected `PHASE_4.2_COMPLETION_REPORT.md`, which described a PDF threading
  refactor that was never applied to the code.
- Replaced `pip install -r requirements.txt` instructions with `uv sync`, and
  removed placeholder GitHub, Discord and support-email links.

---

## [2.0.0] - 2024-12-06

### Major Release - Architectural Refactoring & Performance Optimization

This release represents a complete architectural overhaul with massive performance improvements and code quality enhancements.

### Added
- **Core Modules Architecture**
  - Created `core/excel_utils.py` for centralized file I/O operations
  - Created `core/excel_writer.py` for optimized Excel writing and formatting
  - Created `config/constants.py` for shared configuration values

- **Documentation Suite**
  - Comprehensive README.md with project overview and quick start
  - ARCHITECTURE.md detailing system design and technical architecture
  - USER_GUIDE.md with step-by-step usage instructions
  - DEVELOPER_GUIDE.md for contributors and developers
  - CHANGELOG.md for version history tracking

- **Performance Optimizations (Phase 1-3)**
  - Replaced `df.iterrows()` with `dataframe_to_rows()` (100-1000x faster)
  - Implemented vectorized pandas string operations (10-50x faster)
  - Removed exception handling from hot loops (10-100x faster)
  - Single DataFrame copy pattern (50% less memory)
  - In-memory Excel formatting (50% less I/O)
  - Background threading for Excel Formatter (non-blocking UI)

- **UI Enhancements**
  - Modern Office-style theme with Excel green branding
  - Responsive tabbed interface
  - Cancel button for long-running operations (Excel Formatter)
  - Real-time progress updates and logging

### Changed
- **Excel Formatter Tool**
  - Refactored to use `FileProcessorThread` for background processing
  - Consolidated all transformations into single `apply_all_transformations()` pipeline
  - Updated to import shared utilities from `core/` modules
  - Improved memory efficiency with single-pass transformations

- **Excel Merge & Split Tool**
  - Updated to import shared utilities from `core/` modules
  - Removed ~90 lines of duplicate code

- **PDF to Excel Tool**
  - Maintained existing functionality
  - Ready for Phase 4.2 threading enhancement

- **Desktop Suite App**
  - Window size increased to 1050x650 for better usability
  - Added QScrollArea wrappers for all tools to handle content overflow
  - Centered window positioning on launch

### Removed
- **Code Duplication Eliminated**
  - Removed ~260 lines of duplicate utility functions across tool files
  - Centralized `list_excel_files_in_folder()`, `read_file_to_df()`, `read_all_sheets()`
  - Centralized `save_df_to_excel()` and formatting functions

### Fixed
- Window sizing issues causing tools to be truncated
- Memory leaks from multiple DataFrame copies
- UI freezing during file processing
- Inconsistent date format handling
- Number conversion accuracy issues

### Performance Benchmarks
| Operation | Before | After | Improvement |
|-----------|--------|-------|-------------|
| 10 files (10K rows) | 45s | 2s | **22.5x faster** |
| 50 files (10K rows) | 4min | 10s | **24x faster** |
| Memory usage | 400MB | 200MB | **50% reduction** |

---

## [1.1.0] - 2024-11-15

### Added
- Excel Merge & Split tool
  - Merge multiple Excel/CSV files into single workbook
  - Split workbooks into separate files
  - Batch folder processing

- Excel Formatter improvements
  - Text case transformation (UPPER, lower, Title Case)
  - Currency symbol customization for number formatting
  - Remove duplicate rows functionality

### Changed
- Updated PyQt5 dependency to 5.15+
- Improved error handling in PDF extraction

### Fixed
- PDF table detection accuracy issues
- CSV encoding problems with special characters
- Excel formula preservation during formatting

---

## [1.0.0] - 2024-10-01

### Initial Release

### Added
- **Desktop Suite Application**
  - Main window with tabbed interface
  - Professional modern UI theme

- **PDF to Excel Tool**
  - Extract tables from PDF documents
  - Multi-table selection and preview
  - Export to formatted Excel workbooks

- **Excel Formatter Tool**
  - Trim whitespace from cells
  - Convert numbers stored as text
  - Normalize date formats
  - Apply number formatting (decimals, currency)
  - Auto-fit column widths
  - Apply Excel themes
  - Batch processing support

- **Core Features**
  - Support for .xlsx, .xls, and .csv files
  - Preview functionality before batch processing
  - Progress bars and logging
  - Error handling and user feedback

### Dependencies
- PyQt5 >= 5.15.0
- pandas >= 1.3.0
- openpyxl >= 3.0.0
- pdfplumber >= 0.7.0
- xlrd >= 2.0.0

---

## [Unreleased]

### Next
- Parse Tally's XML responses in `tally_api/` into DataFrames
- Map the Balance Sheet form onto a fixed Excel template rather than writing a
  key/value dump per page
- Move Balance Sheet state out of the Flask session cookie (~4 KB ceiling) into
  a server-side store
- Custom, user-defined transformation rules
- Dark mode and drag-and-drop file selection in the desktop app

---

## Version History Summary

| Version | Date | Highlights |
|---------|------|------------|
| **2.1.0** | 2026-09-04 | Column-width fix, PDF threading, shared core, web hardening, test suite |
| **2.0.0** | 2024-12-06 | Architectural refactoring, 10-20x performance boost, comprehensive documentation |
| **1.1.0** | 2024-11-15 | Merge & Split tool, enhanced formatting options |
| **1.0.0** | 2024-10-01 | Initial release with PDF extraction and Excel formatting |

---

## Migration Guide

### Upgrading from 1.x to 2.0

**No breaking changes** for end users. The application interface remains the same.

**For developers**:

1. **Import statements changed**:
   ```python
   # Old (1.x)
   from excel_formatter_tool import list_excel_files_in_folder

   # New (2.0)
   from core.excel_utils import list_excel_files_in_folder
   ```

2. **Utility functions moved**:
   - `list_excel_files_in_folder()` -> `core/excel_utils.py`
   - `read_file_to_df()` -> `core/excel_utils.py`
   - `read_all_sheets()` -> `core/excel_utils.py`
   - `save_df_to_excel()` -> `core/excel_writer.py`
   - `apply_formatting_to_workbook()` -> `core/excel_writer.py`

3. **Constants moved**:
   - `SUPPORTED_EXTENSIONS` -> `config/constants.py`

4. **Threading model**:
   - Excel Formatter now uses `FileProcessorThread` (QThread)
   - Other tools still run on main thread (will be updated in 2.1.0)

**Recommended actions**:
- Update any custom scripts that import from tool files
- Review new documentation for architectural changes
- Test custom integrations with new module structure

---

## Deprecation Notices

### Version 2.0.0
- **Deprecated**: `apply_openpyxl_autofit_and_theme()` with file path parameter
  - **Replacement**: Use `apply_formatting_to_workbook()` with Workbook object
  - **Reason**: Eliminates redundant I/O operations (50% faster)
  - **Removal**: Planned for version 3.0.0

### Version 1.1.0
- None

### Version 1.0.0
- Initial release, no deprecations

---

## Credits

### Contributors
- **Core Team**: Architecture design, implementation, documentation
- **Community**: Bug reports, feature requests, testing

### Special Thanks
- PyQt5 team for excellent GUI framework
- pandas community for high-performance data processing
- openpyxl maintainers for Excel file handling
- pdfplumber developers for PDF parsing

---

## Support

- **Documentation**: See [README.md](README.md), [USER_GUIDE.md](USER_GUIDE.md), [ARCHITECTURE.md](ARCHITECTURE.md)

---

**Changelog Maintained By**: CA Office Suite Team
**Format**: [Keep a Changelog](https://keepachangelog.com/)
**Versioning**: [Semantic Versioning](https://semver.org/)
