# CA Firm Office Suite

A comprehensive productivity suite for CA Offices, featuring both a powerful Desktop Suite and a modern Web App Suite developed side-by-side.

## 🚀 Desktop Suite (Offline & Powerful)
Located in the root directory.

*   **Launcher**: `CA_Firm_Office_Suite.bat` (Run this to access all tools)
*   **1. PDF to Excel PRO**: Extracts tables from PDFs (`pdf_to_excel_pro_tool.py`)
*   **2. Excel Formatter**: Cleans, formats, and styles Excel files (`excel_formatter_tool.py`)
*   **4. Excel Merge/Split**: Merges multiple workbooks or splits sheets (`excel_merge_split_tool.py`)

## 🌐 Web App Suite (Online & Accessible)
Located in `flask_app/`.

*   **Launcher**: Option 3 in the main menu.
*   **Current Features**:
    *   PDF table extraction.
*   **Hosting**: Ready for deployment on platforms like Render.
*   **Local Staging**: Uses `uploads/` for temporary storage.

## 🛠 Project Structure
*   `venv/`: Shared Python virtual environment.
*   `legacy_backup/`: Old versions (safe to ignore).
