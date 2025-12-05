@echo off
cd /d "%~dp0"
if not exist venv (
    echo Virtual environment not found. Creating one...
    python -m venv venv
    call venv\Scripts\activate
    pip install -r flask_app\requirements.txt
    pip install PyQt5
) else (
    call venv\Scripts\activate
)
echo Starting PDF to Excel PRO Tool...
start /B pythonw "pdf_to_excel_pro_tool.py"
