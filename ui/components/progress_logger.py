"""
Reusable Progress Bar and Log Display component.

This widget provides a consistent progress bar and timestamped log display
that's used across all tools in the CA Firm Office Suite.
"""

from datetime import datetime
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QProgressBar, QTextEdit, QLabel


class ProgressLogger(QWidget):
    """
    Reusable progress bar and log display widget.

    Combines a progress bar with a timestamped log text area for consistent
    progress reporting across all tools.

    Features:
    - Progress bar with configurable maximum value
    - Timestamped log entries
    - Auto-scrolling to latest log entry
    - Read-only log display
    - Configurable log height

    Example:
        >>> progress_logger = ProgressLogger(log_height=150)
        >>> progress_logger.set_max_progress(100)
        >>> progress_logger.set_progress(50)
        >>> progress_logger.log("Processing file 1 of 10")
        >>> progress_logger.clear()

    Attributes:
        progress (QProgressBar): The progress bar widget
        log_text (QTextEdit): The log display text area
    """

    def __init__(self, log_height=150, show_label=True, parent=None):
        """
        Initialize the ProgressLogger widget.

        Args:
            log_height (int): Maximum height of log text area in pixels
            show_label (bool): Whether to show "Processing Log:" label
            parent (QWidget, optional): Parent widget
        """
        super().__init__(parent)
        self.log_height = log_height
        self.show_label = show_label
        self._build_ui()

    def _build_ui(self):
        """Build the UI components."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)

        # Progress bar
        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        # Optional label
        if self.show_label:
            log_label = QLabel("Processing Log:")
            layout.addWidget(log_label)

        # Log display
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(self.log_height)
        layout.addWidget(self.log_text)

    def set_progress(self, value):
        """
        Set the current progress value.

        Args:
            value (int): Progress value (0 to maximum)
        """
        self.progress.setValue(value)

    def set_max_progress(self, maximum):
        """
        Set the maximum progress value.

        Args:
            maximum (int): Maximum value for progress bar
        """
        self.progress.setMaximum(maximum)

    def log(self, message, timestamp_format="[%H:%M:%S]"):
        """
        Add a timestamped message to the log.

        Args:
            message (str): Message to log
            timestamp_format (str): strftime format string for timestamp
        """
        timestamp = datetime.now().strftime(timestamp_format)
        self.log_text.append(f"{timestamp} {message}")
        # Auto-scroll to bottom
        scrollbar = self.log_text.verticalScrollBar()
        if scrollbar is not None:
            scrollbar.setValue(scrollbar.maximum())

    def clear(self):
        """Clear the log text and reset progress to 0."""
        self.log_text.clear()
        self.progress.setValue(0)

    def clear_log_only(self):
        """Clear only the log text, keep progress value."""
        self.log_text.clear()

    def get_log_text(self):
        """
        Get the full log text content.

        Returns:
            str: All log text
        """
        return self.log_text.toPlainText()
