@echo off
cd /d "%~dp0"
call venv\Scripts\activate
start "" pythonw excel_formatter_tool.py
exit
