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
echo Starting Flask Web Server...
echo Open http://127.0.0.1:5000 in your browser.
python flask_app\app.py
pause
