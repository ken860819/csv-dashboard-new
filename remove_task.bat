@echo off
setlocal
set TASK_NAME=CsvDashboardUpdater
schtasks /Delete /TN "%TASK_NAME%" /F
if %ERRORLEVEL% EQU 0 (
  echo 已移除排程 %TASK_NAME%
) else (
  echo 移除排程失敗
)
endlocal
