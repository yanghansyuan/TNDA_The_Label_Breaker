@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Checking openpyxl...
python -c "from openpyxl import Workbook" 2>nul || (echo Installing openpyxl... & pip install openpyxl)
python "%~dp0merge_score_sheets.py"
pause
