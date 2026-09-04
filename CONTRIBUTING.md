# Contributing to CA Firm Office Suite

Thank you for your interest in contributing to the CA Firm Office Suite! This document provides guidelines and instructions for contributing to the project.

---

## Table of Contents

1. [Code of Conduct](#code-of-conduct)
2. [How Can I Contribute?](#how-can-i-contribute)
3. [Development Setup](#development-setup)
4. [Pull Request Process](#pull-request-process)
5. [Coding Standards](#coding-standards)
6. [Commit Message Guidelines](#commit-message-guidelines)
7. [Issue Reporting Guidelines](#issue-reporting-guidelines)

---

## Code of Conduct

### Our Pledge

We are committed to providing a welcoming and inspiring community for all. Please be respectful and constructive in your interactions.

### Expected Behavior

- Be respectful and inclusive
- Welcome newcomers and help them get started
- Accept constructive criticism gracefully
- Focus on what's best for the community
- Show empathy towards other community members

### Unacceptable Behavior

- L Harassment, discrimination, or offensive comments
- L Trolling, insulting/derogatory comments
- L Public or private harassment
- L Publishing others' private information
- L Other conduct deemed inappropriate

---

## How Can I Contribute?

### Reporting Bugs

**Before submitting a bug report**:
2. Try the latest version to see if the issue is fixed
3. Collect information about the bug

**Bug Report Template**:
```markdown
**Describe the bug**
A clear description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:
1. Go to '...'
2. Click on '...'
3. See error

**Expected behavior**
What you expected to happen.

**Screenshots**
If applicable, add screenshots.

**Environment:**
 - OS: [e.g. Windows 10]
 - Python Version: [e.g. 3.9.7]
 - Application Version: [e.g. 2.0.0]

**Additional context**
Any other information about the problem.
```

---

### Suggesting Enhancements

**Enhancement Suggestion Template**:
```markdown
**Is your feature request related to a problem?**
A clear description of the problem.

**Describe the solution you'd like**
What you want to happen.

**Describe alternatives you've considered**
Other approaches you've thought about.

**Additional context**
Screenshots, mockups, or examples.
```

---

### Contributing Code

**Types of contributions we're looking for**:
- = Bug fixes
- ( New features
- Documentation improvements
- Performance optimizations
- Test coverage improvements
- UI/UX enhancements

---

## Development Setup

### 1. Fork and Clone

```bash
# Fork the repository on GitHub, then:
git clone https://github.com/YOUR_USERNAME/ca-office-suite.git
cd ca-office-suite

# Add upstream remote
git remote add upstream https://github.com/original-repo/ca-office-suite.git
```

### 2. Create Virtual Environment

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies

```bash
# Installs runtime and dev dependencies from pyproject.toml, pinned by uv.lock
uv sync
```

### 4. Create Feature Branch

```bash
git checkout -b feature/your-feature-name
```

**Branch naming conventions**:
- `feature/description` - New features
- `bugfix/description` - Bug fixes
- `docs/description` - Documentation updates
- `refactor/description` - Code refactoring
- `test/description` - Test additions/improvements

---

## Pull Request Process

### Before Submitting

1. **Ensure your code works**:
   ```bash
   python desktop_suite_app.py
   ```

2. **Run code quality checks**:
   ```bash
   # Format code and sort imports
   uv run ruff format .

   # Lint code (add --fix to apply safe fixes)
   uv run ruff check .

   # Type checking
   uv run mypy .
   ```

3. **Run tests** (when available):
   ```bash
   pytest
   pytest --cov=core --cov=excel_formatter_tool
   ```

4. **Update documentation**:
   - Update USER_GUIDE.md if adding user-facing features
   - Update ARCHITECTURE.md if changing architecture
   - Update CHANGELOG.md with your changes
   - Add/update docstrings for new functions

---

### Submitting the Pull Request

1. **Commit your changes**:
   ```bash
   git add .
   git commit -m "Add feature: description"
   ```

2. **Push to your fork**:
   ```bash
   git push origin feature/your-feature-name
   ```

3. **Create Pull Request on GitHub**:
   - Go to your fork on GitHub
   - Click "New Pull Request"
   - Select your branch
   - Fill out the PR template

**Pull Request Template**:
```markdown
## Description
Brief description of changes.

## Type of Change
- [ ] Bug fix (non-breaking change fixing an issue)
- [ ] New feature (non-breaking change adding functionality)
- [ ] Breaking change (fix or feature causing existing functionality to change)
- [ ] Documentation update

## How Has This Been Tested?
Describe testing performed.

## Checklist:
- [ ] Code follows project style guidelines
- [ ] Self-review performed
- [ ] Code commented where necessary
- [ ] Documentation updated
- [ ] No new warnings generated
- [ ] Tests added/updated
- [ ] All tests pass
- [ ] CHANGELOG.md updated
```

---

### Review Process

1. **Automated Checks**: CI/CD runs tests and linting
2. **Code Review**: Maintainer reviews your code
3. **Feedback**: Address any requested changes
4. **Approval**: Maintainer approves PR
5. **Merge**: PR is merged to main branch

**What reviewers look for**:
- Code quality and style
- Test coverage
- Documentation completeness
- Performance impact
- Security considerations
- Backward compatibility

---

## Coding Standards

### Python Style Guide

Follow **PEP 8** with these additions:

#### Line Length
```python
# Maximum 120 characters (configured under [tool.ruff] in pyproject.toml)
# Preferred 80-100 for readability
```

#### Naming Conventions
```python
# Classes: PascalCase
class ExcelFormatter:
    pass


# Functions/methods: snake_case
def process_file(path):
    pass


# Constants: UPPER_SNAKE_CASE
MAX_FILE_SIZE = 1024 * 1024


# Private: _leading_underscore
def _internal_helper():
    pass
```

#### Imports
```python
# Standard library
import os
import sys

# Third-party
import pandas as pd
from PyQt5.QtWidgets import QWidget

# Local
from core.excel_utils import read_file_to_df
```

#### Docstrings
```python
def function_name(param1: str, param2: int) -> bool:
    """Brief description.

    Detailed description if needed.

    Args:
        param1: Description of param1
        param2: Description of param2

    Returns:
        Description of return value

    Raises:
        ValueError: When param1 is invalid

    Example:
        >>> function_name("test", 5)
        True
    """
    pass
```

---

### Code Quality Tools

Two tools cover everything. Both read their settings from `pyproject.toml`,
and the same settings drive the VS Code extensions, so the editor and the
terminal always agree.

#### Ruff (Formatter, Import Sorter and Linter)
```bash
# Format all files and sort imports
uv run ruff format .

# Check formatting without modifying
uv run ruff format --check .

# Lint all files
uv run ruff check .

# Apply the safe automatic fixes
uv run ruff check --fix .

# Lint a specific file
uv run ruff check excel_formatter_tool.py
```

To silence a finding on one line, add `# noqa: <CODE>` with a short reason.
Ruff reports a `noqa` that no longer suppresses anything, so stale ones do not
accumulate.

#### mypy (Type Checker)
```bash
# Type check the whole project
uv run mypy .

# Type check a specific file
uv run mypy excel_formatter_tool.py
```

---

## Commit Message Guidelines

### Format
```
<type>(<scope>): <subject>

<body>

<footer>
```

### Types
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes (formatting, etc.)
- `refactor`: Code refactoring
- `perf`: Performance improvements
- `test`: Test additions/changes
- `chore`: Build process, dependencies, etc.

### Examples

**Feature**:
```
feat(formatter): add remove special characters option

- Added checkbox to UI
- Implemented vectorized removal using regex
- Updated documentation
- Added unit tests

Closes #123
```

**Bug Fix**:
```
fix(pdf): handle PDFs with no detectable tables

Previously crashed with IndexError when PDF had no tables.
Now shows user-friendly message instead.

Fixes #456
```

**Documentation**:
```
docs(readme): update installation instructions

- Added virtual environment setup
- Clarified Python version requirements
- Added troubleshooting section
```

---

## Issue Reporting Guidelines

### Bug Reports

**Required Information**:
1. **Environment**: OS, Python version, app version
2. **Steps to Reproduce**: Detailed steps
3. **Expected Behavior**: What should happen
4. **Actual Behavior**: What actually happens
5. **Screenshots**: If applicable
6. **Error Messages**: Full traceback if available

**Labels**:
- `bug`: Something isn't working
- `critical`: Severe impact on users
- `high-priority`: Should be fixed soon
- `low-priority`: Nice to fix eventually

---

### Feature Requests

**Required Information**:
1. **Use Case**: Why is this needed?
2. **Proposed Solution**: How should it work?
3. **Alternatives Considered**: Other approaches
4. **Additional Context**: Mockups, examples

**Labels**:
- `enhancement`: New feature or request
- `good-first-issue`: Good for newcomers
- `help-wanted`: Community input needed
- `discussion`: Needs discussion before implementation

---

## Testing Guidelines

### Writing Tests

**Test file structure**:
```
tests/
 unit/
    test_feature.py
 integration/
    test_workflow.py
 fixtures/
     sample_data.xlsx
```

**Test example**:
```python
import pytest
import pandas as pd
from excel_formatter_tool import trim_whitespace


class TestTrimWhitespace:
    """Tests for trim_whitespace function."""

    def test_trim_leading_spaces(self):
        """Test removal of leading spaces."""
        df = pd.DataFrame({"A": ["  hello"]})
        result = trim_whitespace(df)
        assert result["A"][0] == "hello"

    def test_trim_trailing_spaces(self):
        """Test removal of trailing spaces."""
        df = pd.DataFrame({"A": ["hello  "]})
        result = trim_whitespace(df)
        assert result["A"][0] == "hello"

    def test_preserve_inner_spaces(self):
        """Test that inner spaces are preserved."""
        df = pd.DataFrame({"A": ["  hello world  "]})
        result = trim_whitespace(df)
        assert result["A"][0] == "hello world"

    @pytest.mark.parametrize(
        "input,expected",
        [
            ("  test  ", "test"),
            ("no_spaces", "no_spaces"),
            ("", ""),
        ],
    )
    def test_various_inputs(self, input, expected):
        """Test with various input values."""
        df = pd.DataFrame({"A": [input]})
        result = trim_whitespace(df)
        assert result["A"][0] == expected
```

---

## Performance Considerations

### Benchmarking Changes

If your change affects performance:

1. **Benchmark before and after**:
   ```python
   import time


   def benchmark(func, iterations=100):
       start = time.time()
       for _ in range(iterations):
           func()
       return time.time() - start
   ```

2. **Profile with cProfile**:
   ```python
   import cProfile

   cProfile.run("your_function()")
   ```

3. **Report results in PR**:
   ```markdown
   ## Performance Impact
   - Before: 45 seconds
   - After: 2 seconds
   - Improvement: 22.5x faster
   ```

---

## Documentation Standards

### Code Documentation

- **All public functions** must have docstrings
- **Complex logic** should have inline comments
- **Type hints** for function signatures
- **Examples** in docstrings when helpful

### User Documentation

When adding user-facing features:

1. Update **USER_GUIDE.md** with:
   - Step-by-step instructions
   - Screenshots if applicable
   - Common use cases
   - Troubleshooting tips

2. Update **README.md** if needed:
   - Feature list
   - Quick start examples
   - Installation instructions

---

## Getting Help

**Questions?**
- Read the [Developer Guide](DEVELOPER_GUIDE.md)

**First time contributing?**
- Look for issues labeled `good-first-issue`
- Ask questions in issue comments
- Join our Discord for mentorship

---

## Recognition

**Contributors are recognized**:
- Listed in CHANGELOG.md for their contributions
- Mentioned in release notes
- Added to contributors list in README.md

**Types of contributions recognized**:
- Code contributions
- Documentation improvements
- Bug reports
- Feature suggestions
- Community support
- Testing and QA

---

## License

By contributing to CA Firm Office Suite, you agree that your contributions will be licensed under the MIT License.

---

**Thank you for contributing to CA Firm Office Suite!** Your contributions help make this tool better for the entire CA community.
