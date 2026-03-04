@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

:: 讓 Python 輸出為 UTF-8，避免重導向到 log 時中文變亂碼或 UnicodeEncodeError
set "PYTHONIOENCODING=utf-8"

:: 學生作業根目錄（本 .bat 在 Script 下）
set "ROOT=%~dp0.."
set "SCRIPT=%~dp0"
set "DOC=%ROOT%\Doc"

if not exist "%DOC%" mkdir "%DOC%"

:: 參數：無 = 跑 1～8；一個數字 N = 只跑步驟 N；兩個數字 A B = 跑步驟 A～B
set "START=1"
set "END=8"
if not "%~1"=="" set "START=%~1"
if not "%~2"=="" (
  set "END=%~2"
) else (
  if not "%~1"=="" set "END=%~1"
)

:: 檢查範圍合法（1～8）
if %START% LSS 1 set "START=1"
if %END% GTR 8 set "END=8"
if %START% GTR %END% (
  echo 參數錯誤：起始步驟不可大於結束步驟。
  pause
  exit /b 1
)

echo ========================================
echo 執行步驟 %START% ～ %END% （共 8 步）
echo ========================================
echo 開始時間：%date% %time%
echo.

for /L %%i in (%START%,1,%END%) do call :run_step %%i

echo ========================================
echo 完成。結束時間：%date% %time%
echo 各步 log：Doc\log_01 ～ log_08
echo ========================================

:: 有跑過任一步驟就附加紀錄
echo. >> "%DOC%\腳本與執行順序說明.txt"
echo 最後一鍵執行（步驟 %START%-%END%）：%date% %time% >> "%DOC%\腳本與執行順序說明.txt"

pause
exit /b 0

:run_step
if %1==1 (
  echo [1/8] check_missing_ai_links.py ...
  python "%SCRIPT%check_missing_ai_links.py" > "%DOC%\log_01_check_missing_ai_links.txt" 2>&1
  echo   → Doc\log_01_check_missing_ai_links.txt
  echo.
  goto :eof
)
if %1==2 (
  echo [2/8] batch_extract_ai_chats.py （擷取所有人 AI 對話，可能需 15–25 分鐘）...
  python "%SCRIPT%batch_extract_ai_chats.py" > "%DOC%\log_02_batch_extract_ai_chats.txt" 2>&1
  echo   → Doc\log_02_batch_extract_ai_chats.txt
  echo.
  goto :eof
)
if %1==3 (
  echo [3/8] analyze_prompt_and_feedback.py ...
  python "%SCRIPT%analyze_prompt_and_feedback.py" > "%DOC%\log_03_analyze_prompt_feedback.txt" 2>&1
  echo   → Doc\log_03_analyze_prompt_feedback.txt
  echo.
  goto :eof
)
if %1==4 (
  echo [4/8] merge_score_sheets.py ...
  python "%SCRIPT%merge_score_sheets.py" > "%DOC%\log_04_merge_score_sheets.txt" 2>&1
  echo   → Doc\log_04_merge_score_sheets.txt
  echo.
  goto :eof
)
if %1==5 (
  echo [5/8] gen_radar_svg.py ...
  python "%SCRIPT%gen_radar_svg.py" > "%DOC%\log_05_gen_radar_svg.txt" 2>&1
  echo   → Doc\log_05_gen_radar_svg.txt
  echo.
  goto :eof
)
if %1==6 (
  echo [6/8] summarize_scores.py ...
  python "%SCRIPT%summarize_scores.py" > "%DOC%\log_06_summarize_scores.txt" 2>&1
  echo   → Doc\log_06_summarize_scores.txt
  echo.
  goto :eof
)
if %1==7 (
  echo [7/8] build_gallery_data.py ...
  python "%SCRIPT%build_gallery_data.py" > "%DOC%\log_07_build_gallery_data.txt" 2>&1
  echo   → Doc\log_07_build_gallery_data.txt
  echo.
  goto :eof
)
if %1==8 (
  echo [8/8] gen_overview_charts.py ...
  python "%SCRIPT%gen_overview_charts.py" > "%DOC%\log_08_gen_overview_charts.txt" 2>&1
  echo   → Doc\log_08_gen_overview_charts.txt
  echo.
  goto :eof
)
goto :eof
