@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Checking matplotlib...
python -c "import matplotlib.pyplot" 2>nul || (
    echo Installing matplotlib...
    pip install matplotlib
)
echo 正在產出評分總覽圖表（長條圖、圓餅圖）...
python "%~dp0gen_overview_charts.py"
pause
