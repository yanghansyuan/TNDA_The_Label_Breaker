@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM 用法：直接雙擊 = 用腳本預設找一個學生；或傳一個資料夾路徑只跑該生
REM 例：run_extract_ai_chat.bat "盧品安_恐懼"
echo Checking Playwright (needed to get real conversation from Gemini/ChatGPT share links)...
python -c "from playwright.sync_api import sync_playwright" 2>nul || (
    echo Installing Playwright...
    pip install playwright
    playwright install chromium
)
if "%~1"=="" (
    python "%~dp0extract_ai_chat_from_docx.py"
) else (
    python "%~dp0extract_ai_chat_from_docx.py" "%~1"
)
if errorlevel 1 pause
