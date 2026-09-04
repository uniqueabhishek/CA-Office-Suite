"""
Shared pytest fixtures and path setup.

Puts the repository root and flask_app/ on sys.path so tests can import the
shared core/ modules and the Flask app the same way the apps themselves do.
"""

import os
import sys

import pandas as pd
import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FLASK_APP_DIR = os.path.join(REPO_ROOT, "flask_app")

# The repo root must end up ahead of flask_app on sys.path: both directories
# can hold a package of the same name, and the shared one at the root wins.
# Entries are removed before reinserting, since a path already present would
# otherwise keep its old, lower priority position.
for path in (FLASK_APP_DIR, REPO_ROOT):
    if path in sys.path:
        sys.path.remove(path)
    sys.path.insert(0, path)

# Importing flask_app.app raises unless a signing key is configured, which is the
# point of that check: only a direct development run may go without one.
os.environ.setdefault("SECRET_KEY", "test-only-signing-key")


@pytest.fixture
def messy_df():
    """A DataFrame exercising every transformation: padding, text numbers, dates."""
    return pd.DataFrame(
        {
            "Name": ["  Alice  ", "bob smith", "  CAROL"],
            "Amount": ["1000", "2500.5", "300"],
            "Date": ["01-02-2024", "15-03-2024", "28-12-2023"],
        }
    )


@pytest.fixture
def simple_df():
    """A small DataFrame with short values, for column-width checks."""
    return pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
