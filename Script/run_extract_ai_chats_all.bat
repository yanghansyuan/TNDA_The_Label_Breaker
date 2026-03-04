@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Checking Playwright...
python -c "from playwright.sync_api import sync_playwright" 2>nul || (
    echo Installing Playwright...
    pip install playwright
    playwright install chromium
)
python "%~dp0batch_extract_ai_chats.py"
pause
