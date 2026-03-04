@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Checking pypdf...
python -c "from pypdf import PdfReader" 2>nul || (echo Installing pypdf... & pip install pypdf)
python "%~dp0check_missing_ai_links.py"
pause
