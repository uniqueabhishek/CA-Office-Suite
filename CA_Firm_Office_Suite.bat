@echo off
cd /d "%~dp0"
call venv\Scripts\activate

title CA Firm Office Suite
cls
:menu
cls
echo ==================================================
echo           CA Firm Office Suite
echo ==================================================
echo.
echo    [1] PDF to Excel PRO Tool
echo    [2] Excel Formatter
echo    [3] Web App Server (Local)
echo    [4] Excel Merge/Split Tool
echo    [5] Exit
echo.
echo ==================================================
set /p choice="Enter your choice (1-5): "

if "%choice%"=="1" goto run_pdf
if "%choice%"=="2" goto run_excel
if "%choice%"=="3" goto run_server
if "%choice%"=="4" goto run_merge_split
if "%choice%"=="5" goto exit
goto menu

:run_pdf
start "" "venv\Scripts\pythonw.exe" pdf_to_excel_pro_tool.py
goto menu

:run_excel
start "" "venv\Scripts\pythonw.exe" excel_formatter_tool.py
goto menu

:run_server
echo Starting Flask Server...
start "Flask Server" cmd /k "call venv\Scripts\activate && python flask_app\app.py"
goto menu

:run_merge_split
start "" "venv\Scripts\pythonw.exe" excel_merge_split_tool.py
goto menu

:exit
exit
