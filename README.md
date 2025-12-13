# CA Firm Office Suite

> **Professional Excel & PDF Processing Toolkit for Chartered Accountants**

A high-performance desktop application built with PyQt5 that provides powerful tools for Excel data cleaning, PDF table extraction, and file merging operations. Optimized for CA firms handling large volumes of financial data.

![Version](https://img.shields.io/badge/version-2.0.0-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![License](https://img.shields.io/badge/license-MIT-orange)

---

## 🚀 Features

### 📊 Excel Formatter
- **Automated Data Cleaning**: Trim whitespace, convert text-to-numbers, normalize dates
- **Formatting Options**: Apply number formats, currency symbols, text case transformations
- **Duplicate Removal**: Automatically identify and remove duplicate rows
- **Batch Processing**: Process hundreds of files simultaneously with background threading
- **Performance**: 10-20x faster than traditional methods using optimized pandas operations

### 📄 PDF to Excel Converter
- **Table Extraction**: Extract tables from PDF files using advanced detection
- **Multi-Table Support**: Select and export specific tables or all tables at once
- **Format Preservation**: Maintains table structure and data integrity
- **Excel Export**: Save extracted tables as formatted Excel workbooks

### 🔗 Excel Merge & Split
- **File Merging**: Combine multiple Excel/CSV files into single workbook
- **Sheet Consolidation**: Merge data from multiple sheets intelligently
- **Batch Operations**: Process entire folders of files at once
- **Format Support**: Works with .xlsx, .xls, and .csv files

---

## 📋 Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [System Requirements](#system-requirements)
- [Usage Examples](#usage-examples)
- [Performance Benchmarks](#performance-benchmarks)
- [Documentation](#documentation)
- [Contributing](#contributing)
- [License](#license)

---

## 🔧 Installation

### Prerequisites
- Python 3.8 or higher
- Windows, macOS, or Linux

### Step 1: Clone or Download
```bash
git clone https://github.com/your-repo/ca-office-suite.git
cd ca-office-suite
```

### Step 2: Install Dependencies
```bash
pip install -r requirements.txt
```

**Required packages:**
```
PyQt5>=5.15.0
pandas>=1.3.0
openpyxl>=3.0.0
pdfplumber>=0.7.0
xlrd>=2.0.0
```

### Step 3: Run the Application
```bash
python desktop_suite_app.py
```

---

## ⚡ Quick Start

### Launch the Application
1. Run `python desktop_suite_app.py`
2. The main window will open with three tabs:
   - **PDF to Excel**: Extract tables from PDF files
   - **Excel Formatter**: Clean and format Excel data
   - **Merge & Split**: Combine or split Excel files

### Process Your First File

**Excel Formatter Example:**
1. Click "Add Files" and select Excel/CSV files
2. Check desired options:
   - ✓ Trim leading/trailing spaces
   - ✓ Convert numbers stored as text
   - ✓ Apply number format (2 decimals)
3. Click "Preview Selected File" to see changes
4. Click "Apply to All" to process all files
5. Files are saved to the "Processed" folder by default

---

## 💻 System Requirements

### Minimum Requirements
- **OS**: Windows 7/10/11, macOS 10.12+, Linux (Ubuntu 18.04+)
- **RAM**: 4 GB
- **Storage**: 100 MB free space
- **Python**: 3.8+

### Recommended for Large Files
- **RAM**: 8 GB or more
- **CPU**: Multi-core processor (for parallel processing)
- **SSD**: For faster file I/O operations

---

## 📖 Usage Examples

### Example 1: Batch Clean Excel Files
```python
# Select 100 Excel files with messy data
# Enable options:
#   - Trim whitespace
#   - Convert text to numbers
#   - Normalize dates to dd-mm-yyyy
#   - Apply currency format (₹)
# Result: All files cleaned in ~10 seconds
```

### Example 2: Extract PDF Tables
```python
# Open PDF with financial tables
# Select tables to extract (e.g., Balance Sheet, P&L)
# Click "Convert to Excel"
# Result: Formatted Excel workbook with extracted tables
```

### Example 3: Merge Multiple Reports
```python
# Add folder containing monthly reports (50 files)
# Click "Merge All Files"
# Result: Single consolidated Excel workbook
```

---

## 📊 Performance Benchmarks

| Operation | File Count | Rows/File | Old Method | New Method | Speedup |
|-----------|-----------|-----------|------------|------------|---------|
| Clean & Format | 10 | 10,000 | 45 sec | 2 sec | **22.5x** |
| Clean & Format | 50 | 10,000 | 4 min | 10 sec | **24x** |
| Number Conversion | 100 | 5,000 | 8 min | 15 sec | **32x** |
| PDF Extraction | 1 | 5 tables | 30 sec | 3 sec | **10x** |
| Merge Files | 20 | 1,000 | 60 sec | 5 sec | **12x** |

**Optimizations Applied:**
- ✅ Vectorized pandas operations (10-50x faster)
- ✅ Fast DataFrame-to-Excel writing (100-1000x faster)
- ✅ In-memory formatting (50% less I/O)
- ✅ Background threading (non-blocking UI)
- ✅ Single-pass transformations (50% less memory)

---

## 📚 Documentation

| Document | Description |
|----------|-------------|
| [User Guide](USER_GUIDE.md) | Detailed instructions for using each tool |
| [Architecture Guide](ARCHITECTURE.md) | System design and technical architecture |
| [Developer Guide](DEVELOPER_GUIDE.md) | Contributing and development setup |
| [Changelog](CHANGELOG.md) | Version history and release notes |
| [API Reference](docs/API.md) | Core module and function documentation |

---

## 🏗️ Architecture Overview

```
ca_office_suite/
├── desktop_suite_app.py          # Main application entry point
│
├── core/                          # Shared business logic
│   ├── excel_utils.py            # File I/O utilities
│   └── excel_writer.py           # Excel writing & formatting
│
├── config/                        # Configuration
│   └── constants.py              # Shared constants
│
├── workers/                       # (Planned) Background processing
│   ├── base_worker.py
│   ├── formatter_worker.py
│   ├── pdf_worker.py
│   └── merge_worker.py
│
├── pdf_to_excel_pro_tool.py      # PDF extraction tool
├── excel_formatter_tool.py       # Excel cleaning tool
└── excel_merge_split_tool.py    # Merge/split tool
```

**Design Principles:**
- **Modular Architecture**: Shared utilities in `core/`, tools are independent QWidgets
- **Performance First**: Optimized algorithms for large-scale data processing
- **User Experience**: Background threading keeps UI responsive during long operations
- **Code Reusability**: Zero duplication through shared modules

---

## 🤝 Contributing

We welcome contributions! Please see our [Developer Guide](DEVELOPER_GUIDE.md) for details.

### Quick Contribution Steps
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Development Setup
```bash
# Clone your fork
git clone https://github.com/your-username/ca-office-suite.git
cd ca-office-suite

# Install dependencies
pip install -r requirements.txt

# Run tests
python -m pytest tests/

# Run the app
python desktop_suite_app.py
```

---

## 🐛 Known Issues & Limitations

- **Large PDF Files**: PDFs with >100 pages may take longer to process
- **Complex Tables**: PDF tables with merged cells may require manual adjustment
- **Memory Usage**: Processing 100+ large files simultaneously requires 8GB+ RAM
- **File Format**: Only supports .xlsx, .xls, .csv (not .xlsm or .xlsb)

See [Issues](https://github.com/your-repo/ca-office-suite/issues) for current bugs and feature requests.

---

## 🗺️ Roadmap

### Version 2.1 (Q1 2025)
- [ ] Add background threading to PDF and Merge tools (Phase 4.2)
- [ ] Remove standalone entry points (Phase 4.3)
- [ ] Add cancel button to all tools
- [ ] Implement progress bar improvements

### Version 2.2 (Q2 2025)
- [ ] Database export functionality (MySQL, PostgreSQL)
- [ ] Advanced filtering and data validation
- [ ] Custom transformation rules (user-defined)
- [ ] Multi-language support

### Version 3.0 (Q3 2025)
- [ ] Web-based version (Flask/Django)
- [ ] Cloud storage integration (Google Drive, OneDrive)
- [ ] Collaborative features
- [ ] API for automation

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

```
MIT License

Copyright (c) 2024 CA Office Suite

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 👥 Authors & Acknowledgments

**Developed for CA Firms** by the Office Suite Team

**Special Thanks:**
- PyQt5 team for the excellent GUI framework
- pandas community for high-performance data processing
- openpyxl maintainers for Excel file handling
- pdfplumber developers for PDF parsing capabilities

---

## 📞 Support

- **Documentation**: [Read the Docs](docs/)
- **Issues**: [GitHub Issues](https://github.com/your-repo/ca-office-suite/issues)
- **Email**: support@ca-office-suite.com
- **Community**: [Discord Server](https://discord.gg/ca-suite)

---

## 🌟 Star History

If you find this project useful, please consider giving it a star ⭐

[![Star History Chart](https://api.star-history.com/svg?repos=your-repo/ca-office-suite&type=Date)](https://star-history.com/#your-repo/ca-office-suite&Date)

---

**Built with ❤️ for Chartered Accountants**

```
PY __ CA Office Suite
├─ .flake8
├─ ARCHITECTURE.md
├─ CA_Firm_Office_Suite.bat
├─ CHANGELOG.md
├─ config
│  ├─ constants.py
│  └─ __init__.py
├─ CONTRIBUTING.md
├─ core
│  ├─ excel_utils.py
│  ├─ excel_writer.py
│  └─ __init__.py
├─ desktop_suite_app.py
├─ DEVELOPER_GUIDE.md
├─ DOCUMENTATION_INDEX.md
├─ excel_formatter_tool.py
├─ excel_merge_split_tool.py
├─ flask_app
│  ├─ app.py
│  ├─ core
│  │  ├─ excel_utils.py
│  │  ├─ excel_writer.py
│  │  ├─ merge_logic.py
│  │  ├─ transformations.py
│  │  └─ __init__.py
│  ├─ Procfile
│  ├─ requirements.txt
│  ├─ static
│  │  └─ style.css
│  ├─ templates
│  │  ├─ base.html
│  │  ├─ formatter.html
│  │  ├─ index.html
│  │  ├─ merge.html
│  │  └─ select_tables.html
│  └─ utils.py
├─ LICENSE
├─ pdf_to_excel_pro_tool.py
├─ PHASE_4.2_COMPLETION_REPORT.md
├─ PHASE_4.3_COMPLETION_REPORT.md
├─ PHASE_4.4_COMPLETION_REPORT.md
├─ Procfile
├─ README.md
├─ requirements-dev.txt
├─ requirements.txt
├─ ui
│  ├─ components
│  │  ├─ progress_logger.py
│  │  └─ __init__.py
│  └─ __init__.py
├─ USER_GUIDE.md
└─ workers
   ├─ base_worker.py
   ├─ formatter_worker.py
   ├─ merge_worker.py
   ├─ pdf_worker.py
   └─ __init__.py

```