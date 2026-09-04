"""
Shared constants and configuration values for CA Firm Office Suite.

This module centralizes all configuration values used across the application
to ensure consistency and ease of maintenance.
"""

# File format support
SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv")

# Default settings
DEFAULT_OUTPUT_FOLDER = "Processed"
DEFAULT_CURRENCY_SYMBOL = "₹"
DEFAULT_DATE_FORMAT = "dd-mm-yyyy"

# UI settings
WINDOW_WIDTH = 1050
WINDOW_HEIGHT = 650

# Processing settings
MAX_PREVIEW_ROWS = 50

# Generated worksheet column widths (in characters).
# Widths are sized to the longest cell value, then clamped to this range.
MIN_COLUMN_WIDTH = 10
MAX_COLUMN_WIDTH = 60

# Excel hard limit on worksheet name length
MAX_SHEET_NAME_LENGTH = 31

# Web upload settings
ALLOWED_UPLOAD_EXTENSIONS = (".xlsx", ".xls", ".csv")
UPLOAD_RETENTION_SECONDS = 60 * 60  # Delete stray uploads older than 1 hour
