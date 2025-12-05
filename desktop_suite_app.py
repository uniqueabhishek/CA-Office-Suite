import sys
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QTabWidget,
    QVBoxLayout,
    QWidget,
    QLabel,
    QStyleFactory,
    QDesktopWidget
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
    background-color: #f0f2f5; /* Light Gray Background */
}
QWidget {
    font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
    font-size: 14px;
    color: #202020;
}

/* Header Styling - Premium Brand Look */
QLabel#HeaderLabel {
    background-color: #107c41; /* Excel Green Brand Color */
    color: #ffffff;
    font-size: 24px;
    font-weight: 600;
    padding: 18px 30px;
    border-bottom: 4px solid #0c5c30; /* Darker accent border */
}

/* Tab Widget Styling - Card Look */
QTabWidget::pane {
    border: 1px solid #e0e0e0;
    background: #ffffff;
    border-radius: 8px; /* Softer corners */
    border-top-left-radius: 0px; /* Connects to active tab */
    margin-top: -1px;
    /* Subtle shadow effect simulated with border */
    border-bottom: 2px solid #d0d0d0;
    border-right: 2px solid #d0d0d0;
}

QTabBar::tab {
    background: #e5e5e5;
    border: none;
    padding: 12px 28px;
    margin-right: 2px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    color: #555555;
    font-weight: 500;
    min-width: 130px;
}

QTabBar::tab:selected {
    background: #ffffff;
    color: #107c41; /* Brand Green */
    font-weight: bold;
    border-bottom: 1px solid #ffffff; /* Merge with pane */
    border-top: 3px solid #107c41; /* Top Brand Line */
}

QTabBar::tab:hover:!selected {
    background: #dcdcdc;
    color: #000000;
}

/* GroupBox Styling */
QGroupBox {
    border: 1px solid #e0e0e0;
    border-radius: 6px;
    margin-top: 1.2em;
    padding-top: 10px;
    background-color: #fafafa;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 10px;
    color: #107c41;
    font-weight: 700;
}

/* Input Fields */
QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox {
    border: 1px solid #c8c8c8;
    border-radius: 4px;
    padding: 8px 10px;
    background-color: #ffffff;
    selection-background-color: #107c41;
    selection-color: white;
}
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {
    border: 2px solid #107c41;
    background-color: #ffffff;
}

/* Combo Box */
QComboBox {
    border: 1px solid #c8c8c8;
    border-radius: 4px;
    padding: 6px 12px;
    min-width: 6em;
    background: #ffffff;
}
QComboBox:hover {
    border: 1px solid #107c41;
}
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 20px;
    border-left-width: 0px;
}

/* Buttons - Primary Action Look */
QPushButton {
    background-color: #107c41;
    color: white;
    border: 1px solid #107c41;
    padding: 9px 22px;
    border-radius: 4px;
    font-weight: 600;
    font-size: 14px;
}
QPushButton:hover {
    background-color: #159e52;
    border-color: #159e52;
}
QPushButton:pressed {
    background-color: #0c5c30;
    border-color: #0c5c30;
}
QPushButton:disabled {
    background-color: #e0e0e0;
    border-color: #e0e0e0;
    color: #888888;
}

/* Tables - Data Grid Look */
QTableWidget {
    background-color: #ffffff;
    alternate-background-color: #fefefe;
    gridline-color: #e0e0e0;
    border: 1px solid #d0d0d0;
    selection-background-color: #e6f7ec;
    selection-color: #000000;
    font-size: 13px;
}
QHeaderView::section {
    background-color: #f9f9f9;
    padding: 10px;
    border: none;
    border-bottom: 2px solid #107c41;
    border-right: 1px solid #e0e0e0;
    font-weight: 700;
    color: #444444;
    text-transform: uppercase;
    font-size: 12px;
}

/* Scrollbars - Minimalist */
QScrollBar:vertical {
    border: none;
    background: #f1f1f1;
    width: 14px;
    margin: 0px;
}
QScrollBar::handle:vertical {
    background: #c1c1c1;
    min-height: 20px;
    border-radius: 7px;
    margin: 2px;
}
QScrollBar::handle:vertical:hover {
    background: #a8a8a8;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

/* Progress Bar */
QProgressBar {
    border: none;
    background-color: #e0e0e0;
    border-radius: 4px;
    text-align: center;
    color: white;
    font-weight: bold;
}
QProgressBar::chunk {
    background-color: #107c41;
    border-radius: 4px;
}
"""


class DesktopSuiteApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CA Firm Office Suite - Desktop")
        self.resize(800, 600)  # Reduced size for better laptop compatibility

        self.center_window()

        # Setup Theme
        QApplication.setStyle(QStyleFactory.create("Fusion"))
        self.setStyleSheet(MODERN_STYLESHEET)

        self.setup_ui()

    def center_window(self):
        # Center the window on the screen
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) // 2,
                  (screen.height() - size.height()) // 2)

    def setup_ui(self):
        # Main Layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(0, 0, 0, 0)  # Edge to edge
        layout.setSpacing(10)

        # Header
        header = QLabel("CA Firm Office Suite")
        header.setObjectName("HeaderLabel")  # For CSS targeting
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
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
