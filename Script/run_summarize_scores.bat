@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在彙整所有學生評分表並產出總覽...
python "%~dp0summarize_scores.py"
pause
