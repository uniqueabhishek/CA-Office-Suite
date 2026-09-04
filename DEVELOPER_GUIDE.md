# Developer Guide - CA Firm Office Suite

> **Contributing, Development Setup, and Technical Guidelines**

This guide is for developers who want to contribute to the CA Office Suite project or understand the codebase for extension and customization.

---

## Table of Contents

1. [Development Environment Setup](#development-environment-setup)
2. [Project Structure](#project-structure)
3. [Coding Standards](#coding-standards)
4. [Adding New Features](#adding-new-features)
5. [Testing Guidelines](#testing-guidelines)
6. [Performance Optimization](#performance-optimization)
7. [Debugging Tips](#debugging-tips)
8. [Contributing Workflow](#contributing-workflow)
9. [Release Process](#release-process)

---

## Development Environment Setup

### Prerequisites
- **Python**: 3.10 or higher (uv installs it for you)
- **Git**: For version control
- **IDE**: VS Code, PyCharm, or similar (VS Code recommended)
- **OS**: Windows, macOS, or Linux

### Step 1: Create Virtual Environment
This project uses [uv](https://docs.astral.sh/uv/) for environment management.
uv downloads its own Python, so no system Python install is required.

```bash
# Create virtual environment (Python 3.10)
uv venv

# Activate virtual environment
# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate
```

### Step 2: Install Dependencies
```bash
uv sync
```

`uv sync` installs both the runtime dependencies and the `dev` dependency
group from `pyproject.toml`, pinned by `uv.lock`. There is no
`requirements.txt`; add or remove packages with `uv add` / `uv remove`.

The dev group in `pyproject.toml` includes:
```
# Testing
pytest>=7.0.0
pytest-qt>=4.0.0
pytest-cov>=3.0.0

# Code Quality
ruff>=0.16.6
mypy>=0.950

# Documentation
sphinx>=4.5.0
```

### Step 3: Verify Installation
```bash
# Run the application
uv run python desktop_suite_app.py

# Run tests
uv run pytest

# Check code style
uv run ruff check .
```

---

## Project Structure

```
ca_office_suite/
|
+- desktop_suite_app.py            # Desktop entry point
|
+- core/                           # Shared logic (desktop + web)
|  +- excel_utils.py               # File I/O, name sanitising
|  +- excel_writer.py              # Excel writing/formatting
|  +- transformations.py           # Cleaning rules
|  +- merge_logic.py               # Merge/concat rules
|
+- config/                         # Configuration
|  +- constants.py                 # Application constants
|
+- workers/                        # QThread background workers
+- ui/components/                  # Reusable widgets
|
+- pdf_to_excel_pro_tool.py        # PDF extraction tool
+- excel_formatter_tool.py         # Excel cleaning tool
+- excel_merge_split_tool.py       # Merge/split tool
|
+- flask_app/                      # Web front-end
|  +- app.py                       # Routes, upload handling
|  +- utils.py                     # Bridge to core/
|  +- blueprints/                  # Balance sheet generator
|
+- tally_api/                      # Tally XML integration
|
+- tests/                          # pytest suite
|  +- conftest.py
|  +- test_excel_utils.py
|  +- test_excel_writer.py
|  +- test_transformations.py
|  +- test_merge_logic.py
|  +- test_web_utils.py
|  +- test_flask_routes.py
|
+- pyproject.toml                  # Dependencies, ruff, mypy and pytest config
+- uv.lock                         # Pinned dependency versions
```

---

## Coding Standards

### Python Style Guide
Follow **PEP 8** with these specific guidelines:

#### Naming Conventions
```python
# Classes: PascalCase
class ExcelCleanerWindow(QWidget):
    pass


# Functions/methods: snake_case
def read_file_to_df(path):
    pass


# Constants: UPPER_SNAKE_CASE
SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv")


# Private methods: _leading_underscore
def _build_ui(self):
    pass
```

#### Import Organization
```python
# Standard library imports
import os
import sys
from datetime import datetime

# Third-party imports
import pandas as pd
from PyQt5.QtWidgets import QWidget, QVBoxLayout

# Local application imports
from core.excel_utils import read_file_to_df
from config.constants import SUPPORTED_EXTENSIONS
```

#### Docstrings
Use **Google-style** docstrings:

```python
def save_df_to_excel(df, path, sheet_name="Sheet1", number_format_map=None):
    """Save a DataFrame to an Excel file with formatting.

    Args:
        df (pd.DataFrame): DataFrame to save
        path (str): Output file path
        sheet_name (str, optional): Sheet name. Defaults to "Sheet1".
        number_format_map (dict, optional): Column format mapping.

    Returns:
        None

    Raises:
        IOError: If file cannot be written
        ValueError: If DataFrame is empty

    Example:
        >>> df = pd.DataFrame({"A": [1, 2, 3]})
        >>> save_df_to_excel(df, "output.xlsx")
    """
    pass
```

#### Type Hints
Use type hints for function signatures:

```python
from typing import List, Dict, Optional
import pandas as pd


def read_all_sheets(path: str) -> Dict[str, pd.DataFrame]:
    """Read all sheets from an Excel file."""
    pass


def process_files(files: List[str], options: Dict[str, bool]) -> Optional[str]:
    """Process files with given options."""
    pass
```

---

### Code Formatting

#### Line Length
- **Maximum**: 120 characters (configured under `[tool.ruff]` in `pyproject.toml`)
- **Preferred**: 80-100 characters for readability

#### Blank Lines
```python
# Two blank lines between top-level definitions
class MyClass:
    pass


def my_function():
    pass


# One blank line between methods
class MyClass:
    def method_one(self):
        pass

    def method_two(self):
        pass
```

#### String Quotes
- **Prefer double quotes** `"` for strings
- **Use single quotes** `'` for dict keys and short literals
- **Use triple double quotes** `"""` for docstrings

---

## Adding New Features

### Feature Development Workflow

#### 1. Plan the Feature
- Create GitHub issue describing the feature
- Discuss design approach
- Get approval before implementing

#### 2. Create Feature Branch
```bash
git checkout -b feature/feature-name
```

#### 3. Implement the Feature

**For new tool**:
1. Create new file: `new_tool.py`
2. Define QWidget class
3. Implement UI in `_build_ui()` method
4. Add business logic
5. Connect signals/slots

**For new utility function**:
1. Add to appropriate `core/` module
2. Write docstring
3. Add type hints
4. Export in `__init__.py`

**For new transformation**:
1. Add function to `core/transformations.py` (shared by desktop and web)
2. Follow single-pass transformation pattern
3. Use vectorized pandas operations
4. Update `apply_all_transformations()` pipeline
5. Add a test in `tests/test_transformations.py`

#### 4. Test the Feature
```bash
# Manual testing
uv run python desktop_suite_app.py

# Automated testing
uv run pytest tests/test_new_feature.py
```

#### 5. Document the Feature
- Update `USER_GUIDE.md` with usage instructions
- Update `ARCHITECTURE.md` if architecture changed
- Add docstrings to new functions
- Update `CHANGELOG.md`

#### 6. Create Pull Request
```bash
git add .
git commit -m "Add feature: description"
git push origin feature/feature-name
```

Then create PR on GitHub.

---

### Example: Adding a New Transformation

Let's add a "Remove Special Characters" transformation to Excel Formatter.

**Step 1**: Add transformation function
```python
# In excel_formatter_tool.py


def remove_special_characters(df):
    """Remove special characters from string columns.

    Args:
        df (pd.DataFrame): Input DataFrame

    Returns:
        pd.DataFrame: DataFrame with special characters removed

    Example:
        >>> df = pd.DataFrame({"A": ["Hello@World!", "Test#123"]})
        >>> result = remove_special_characters(df)
        >>> print(result["A"].tolist())
        ['HelloWorld', 'Test123']
    """
    df2 = df.copy()
    for col in df2.columns:
        if df2[col].dtype == object:
            # Use vectorized str.replace with regex
            df2[col] = df2[col].str.replace(r"[^a-zA-Z0-9\s]", "", regex=True)
    return df2
```

**Step 2**: Add UI control
```python
# In ExcelCleanerWindow._build_ui()

self.chk_remove_special = QCheckBox("Remove special characters")
fix_layout.addWidget(self.chk_remove_special)
```

**Step 3**: Integrate into pipeline
```python
# In apply_all_transformations()

# Add after text case transformation
if options.get("remove_special_chars", False):
    for col in result.columns:
        if result[col].dtype == object:
            result[col] = result[col].str.replace(r"[^a-zA-Z0-9\s]", "", regex=True)
```

**Step 4**: Update preview logic
```python
# In preview_selected()

if self.chk_remove_special.isChecked():
    df_preview = remove_special_characters(df_preview)
```

**Step 5**: Add to options dict
```python
# In apply_to_all()

options = {
    # ... existing options ...
    "remove_special_chars": self.chk_remove_special.isChecked(),
}
```

---

## Testing Guidelines

### Test Structure (Planned)

```
tests/
 unit/
    test_excel_utils.py
    test_excel_writer.py
    test_transformations.py
 integration/
    test_formatter_workflow.py
    test_pdf_workflow.py
 fixtures/
     sample_data.xlsx
     sample_pdf.pdf
```

### Writing Unit Tests

**Example test for `trim_whitespace()` function**:

```python
# tests/unit/test_transformations.py

import pytest
import pandas as pd
from excel_formatter_tool import trim_whitespace


class TestTrimWhitespace:
    def test_trim_leading_spaces(self):
        """Test removal of leading spaces."""
        df = pd.DataFrame({"A": ["  hello", "  world"]})
        result = trim_whitespace(df)
        assert result["A"].tolist() == ["hello", "world"]

    def test_trim_trailing_spaces(self):
        """Test removal of trailing spaces."""
        df = pd.DataFrame({"A": ["hello  ", "world  "]})
        result = trim_whitespace(df)
        assert result["A"].tolist() == ["hello", "world"]

    def test_trim_both_sides(self):
        """Test removal of both leading and trailing spaces."""
        df = pd.DataFrame({"A": ["  hello  ", "  world  "]})
        result = trim_whitespace(df)
        assert result["A"].tolist() == ["hello", "world"]

    def test_preserve_inner_spaces(self):
        """Test that inner spaces are preserved."""
        df = pd.DataFrame({"A": ["  hello world  "]})
        result = trim_whitespace(df)
        assert result["A"].tolist() == ["hello world"]

    def test_non_string_columns_unchanged(self):
        """Test that numeric columns are not affected."""
        df = pd.DataFrame({"A": [1, 2, 3], "B": ["  text  "]})
        result = trim_whitespace(df)
        assert result["A"].tolist() == [1, 2, 3]
        assert result["B"].tolist() == ["text"]
```

### Running Tests

```bash
# Run all tests
pytest

# Run specific test file
pytest tests/unit/test_transformations.py

# Run specific test
pytest tests/unit/test_transformations.py::TestTrimWhitespace::test_trim_leading_spaces

# Run with coverage
pytest --cov=core --cov=excel_formatter_tool

# Generate HTML coverage report
pytest --cov=. --cov-report=html
```

---

## Performance Optimization

### Profiling Code

#### Using cProfile
```python
import cProfile
import pstats


def profile_function():
    # Your code here
    pass


# Profile the function
profiler = cProfile.Profile()
profiler.enable()
profile_function()
profiler.disable()

# Print stats
stats = pstats.Stats(profiler)
stats.sort_stats("cumulative")
stats.print_stats(20)  # Top 20 slowest functions
```

#### Using line_profiler
```bash
# Install
pip install line_profiler

# Add @profile decorator to function
@profile
def slow_function():
    pass

# Run profiler
kernprof -l -v script.py
```

### Optimization Checklist

 **DO**:
- Use vectorized pandas operations (`.str`, `.apply` with NumPy)
- Use `dataframe_to_rows()` instead of `iterrows()`
- Minimize DataFrame copies
- Use in-place operations where safe
- Remove exception handling from hot loops
- Use generators for large iterations

L **DON'T**:
- Use `iterrows()` or `apply(lambda)`  for large DataFrames
- Create unnecessary DataFrame copies
- Use exceptions for control flow
- Read/write files multiple times
- Block the main thread with long operations

### Performance Testing

```python
import time
import pandas as pd


# Benchmark function
def benchmark(func, *args, iterations=100):
    """Benchmark a function's execution time."""
    times = []
    for _ in range(iterations):
        start = time.time()
        func(*args)
        end = time.time()
        times.append(end - start)

    avg_time = sum(times) / len(times)
    print(f"{func.__name__}: {avg_time:.4f}s average")


# Example
df = pd.DataFrame({"A": [" text " for _ in range(10000)]})


# Old method (slow)
def trim_with_apply(df):
    df["A"] = df["A"].apply(lambda x: x.strip())


# New method (fast)
def trim_vectorized(df):
    df["A"] = df["A"].str.strip()


benchmark(trim_with_apply, df.copy())  # ~0.5s
benchmark(trim_vectorized, df.copy())  # ~0.02s  (25x faster!)
```

---

## Debugging Tips

### Using PyQt5 Debug Mode

```python
# Enable Qt debug messages
os.environ["QT_DEBUG_PLUGINS"] = "1"
```

### Print Debugging in Workers

```python
class FileProcessorThread(QThread):
    def run(self):
        # Use signals instead of print()
        self.progress_update.emit(0, f"DEBUG: Processing file {file_path}")
```

### Using Python Debugger

```python
import pdb


def problematic_function():
    # Set breakpoint
    pdb.set_trace()
    # Code execution pauses here
    result = some_operation()
    return result


# Run and interact:
# n - next line
# s - step into
# c - continue
# p variable - print variable
# q - quit
```

### VS Code Debugging

**launch.json**:
```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Python: Desktop Suite",
            "type": "python",
            "request": "launch",
            "program": "${workspaceFolder}/desktop_suite_app.py",
            "console": "integratedTerminal",
            "justMyCode": false
        }
    ]
}
```

---

## Contributing Workflow

### 1. Fork & Clone
```bash
# Fork on GitHub, then:
git clone https://github.com/YOUR_USERNAME/ca-office-suite.git
cd ca-office-suite
git remote add upstream https://github.com/original-repo/ca-office-suite.git
```

### 2. Create Feature Branch
```bash
git checkout -b feature/my-new-feature
```

### 3. Make Changes
- Write code
- Add tests
- Update documentation

### 4. Test Changes
```bash
# Run tests
uv run pytest

# Format code
uv run ruff format .

# Lint code
uv run ruff check .

# Type checking
uv run mypy .
```

### 5. Commit Changes
```bash
git add .
git commit -m "Add feature: description

- Detailed explanation
- List of changes
- Related issue #123"
```

### 6. Push & Create PR
```bash
git push origin feature/my-new-feature
```

Then create Pull Request on GitHub.

### 7. Code Review Process
- Address review comments
- Make requested changes
- Push updates to same branch
- PR automatically updates

### 8. Merge
- Maintainer merges after approval
- Delete feature branch

---

## Release Process

### Version Numbering
Follow **Semantic Versioning** (semver):
- **Major**: Breaking changes (e.g., 1.0.0 ’ 2.0.0)
- **Minor**: New features (e.g., 1.1.0 ’ 1.2.0)
- **Patch**: Bug fixes (e.g., 1.1.1 ’ 1.1.2)

### Release Checklist

1. **Update Version Number**
   - `config/constants.py`: `VERSION = "2.1.0"`
   - `README.md`: Badge and version references

2. **Update CHANGELOG.md**
   - Document all changes
   - Categorize: Added, Changed, Fixed, Removed

3. **Run Full Test Suite**
   ```bash
   uv run pytest
   uv run ruff format --check .
   uv run ruff check .
   uv run mypy .
   ```

4. **Update Documentation**
   - README.md
   - USER_GUIDE.md
   - ARCHITECTURE.md

5. **Create Git Tag**
   ```bash
   git tag -a v2.1.0 -m "Release version 2.1.0"
   git push origin v2.1.0
   ```

6. **Build Distribution** (if applicable)
   ```bash
   pyinstaller --onefile --windowed desktop_suite_app.py
   ```

7. **Create GitHub Release**
   - Upload binary (if applicable)
   - Copy changelog for release notes
   - Mark as latest release

---

## Additional Resources

### Useful Documentation
- [PyQt5 Documentation](https://www.riverbankcomputing.com/static/Docs/PyQt5/)
- [pandas Documentation](https://pandas.pydata.org/docs/)
- [openpyxl Documentation](https://openpyxl.readthedocs.io/)
- [pdfplumber Documentation](https://github.com/jsvine/pdfplumber)

### Related Projects
- [xlsxwriter](https://xlsxwriter.readthedocs.io/) - Alternative Excel writer
- [tabula-py](https://github.com/chezou/tabula-py) - Alternative PDF table extractor
- [streamlit](https://streamlit.io/) - For web-based version

### Community
- GitHub Discussions
- Discord Server (link in README)
- Stack Overflow tag: `ca-office-suite`

---

**Developer Guide Version**: 2.0
**Last Updated**: December 2024
**Maintained By**: CA Office Suite Team
