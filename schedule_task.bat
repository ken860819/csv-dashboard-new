@echo off
setlocal
set TASK_NAME=CsvDashboardUpdater
set BASE_DIR=%~dp0
set UPDATER_EXE=%BASE_DIR%CsvDashboardUpdater\CsvDashboardUpdater.exe

if not exist "%UPDATER_EXE%" (
  echo 找不到 %UPDATER_EXE%
  echo 請確認此批次檔與 CsvDashboardUpdater 資料夾在同一層
  exit /b 1
)

schtasks /Create /SC DAILY /ST 12:00 /TN "%TASK_NAME%" /TR "\"%UPDATER_EXE%\"" /F
if %ERRORLEVEL% EQU 0 (
  echo 已建立排程 %TASK_NAME% (每天 12:00)
) else (
  echo 建立排程失敗
)
endlocal
