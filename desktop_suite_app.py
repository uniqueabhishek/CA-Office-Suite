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
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt

# Import our tools as widgets
from pdf_to_excel_pro_tool import PDFTableExtractor
from excel_formatter_tool import ExcelCleanerWindow
from excel_merge_split_tool import ExcelMergeSplitWindow

class DesktopSuiteApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CA Firm Office Suite - Desktop")
        self.resize(1200, 800)

        # Setup Theme
        QApplication.setStyle(QStyleFactory.create("Fusion"))
        self.setup_ui()

    def setup_ui(self):
        # Main Layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # Header
        header = QLabel("CA Firm Office Suite")
        header.setStyleSheet("font-size: 24px; font-weight: bold; margin: 10px; color: #2c3e50;")
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)

        # Tabs
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #C2C7CB; top: -1px; }
            QTabBar::tab {
                background: #E1E1E1;
                border: 1px solid #C4C4C3;
                padding: 10px;
                min-width: 150px;
                font-size: 14px;
            }
            QTabBar::tab:selected { background: #ffffff; margin-bottom: -1px; }
        """)

        # Add Tools as Tabs
        self.pdf_tool = PDFTableExtractor()
        self.formatter_tool = ExcelCleanerWindow()
        self.merge_split_tool = ExcelMergeSplitWindow()

        self.tabs.addTab(self.pdf_tool, "PDF to Excel")
        self.tabs.addTab(self.formatter_tool, "Excel Formatter")
        self.tabs.addTab(self.merge_split_tool, "Merge & Split")

        layout.addWidget(self.tabs)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Global Styling for consistency
    app.setFont(QFont("Segoe UI", 10))

    window = DesktopSuiteApp()
    window.show()
    sys.exit(app.exec_())
