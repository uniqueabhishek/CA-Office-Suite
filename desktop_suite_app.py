import sys
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QTabWidget,
    QVBoxLayout,
    QWidget,
    QLabel,
    QStyleFactory
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

# Import our tools as widgets
from pdf_to_excel_pro_tool import PDFTableExtractor
from excel_formatter_tool import ExcelCleanerWindow
from excel_merge_split_tool import ExcelMergeSplitWindow

# Modern Office Professional Theme
MODERN_STYLESHEET = """
/* Global Application Styling */
QMainWindow {
    background-color: #f3f3f3;
}
QWidget {
    font-family: "Segoe UI", sans-serif;
    color: #333333;
}

/* Header Styling */
QLabel#HeaderLabel {
    color: #217346; /* Excel Green */
    font-size: 26px;
    font-weight: bold;
    padding: 15px;
    background-color: #ffffff;
    border-bottom: 2px solid #217346;
}

/* Tab Widget Styling */
QTabWidget::pane {
    border: 1px solid #dcdcdc;
    background: #ffffff;
    border-radius: 4px;
    margin-top: -1px;
}

QTabBar::tab {
    background: #e1e1e1;
    border: 1px solid #dcdcdc;
    padding: 10px 20px;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    color: #555555;
    font-weight: 500;
    min-width: 120px;
}

QTabBar::tab:selected {
    background: #ffffff;
    border-bottom: 1px solid #ffffff; /* Merge with pane */
    color: #217346; /* Excel Green */
    font-weight: bold;
    border-top: 3px solid #217346; /* Top Highlight */
}

QTabBar::tab:hover:!selected {
    background: #eaeaea;
    color: #333333;
}

/* Button Generic Styling (for consistency across tools if they inherit) */
QPushButton {
    background-color: #217346; /* Excel Green */
    color: white;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #2a9155;
}

QPushButton:pressed {
    background-color: #1a5c38;
}

/* Scrollbars */
QScrollBar:vertical {
    border: none;
    background: #f0f0f0;
    width: 10px;
    margin: 0px 0px 0px 0px;
}
QScrollBar::handle:vertical {
    background: #c1c1c1;
    min-height: 20px;
    border-radius: 5px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    border: none;
    background: none;
}
"""


class DesktopSuiteApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CA Firm Office Suite - Desktop")
        self.resize(1280, 850)

        # Setup Theme
        QApplication.setStyle(QStyleFactory.create("Fusion"))
        self.setStyleSheet(MODERN_STYLESHEET)

        self.setup_ui()

    def setup_ui(self):
        # Main Layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(0, 0, 0, 0) # Edge to edge
        layout.setSpacing(10)

        # Header
        header = QLabel("CA Firm Office Suite")
        header.setObjectName("HeaderLabel") # For CSS targeting
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)

        # Container for content with some margin
        content_container = QWidget()
        content_layout = QVBoxLayout(content_container)
        content_layout.setContentsMargins(20, 10, 20, 20)

        # Tabs
        self.tabs = QTabWidget()
        # Ensure tabs fill the space
        self.tabs.setDocumentMode(False)

        # Add Tools as Tabs
        self.pdf_tool = PDFTableExtractor()
        self.formatter_tool = ExcelCleanerWindow()
        self.merge_split_tool = ExcelMergeSplitWindow()

        self.tabs.addTab(self.pdf_tool, "PDF to Excel")
        self.tabs.addTab(self.formatter_tool, "Excel Formatter")
        self.tabs.addTab(self.merge_split_tool, "Merge & Split")

        content_layout.addWidget(self.tabs)
        layout.addWidget(content_container)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Global Styling for consistency
    app.setFont(QFont("Segoe UI", 10))

    window = DesktopSuiteApp()
    window.show()
    sys.exit(app.exec_())
