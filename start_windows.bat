@echo off
cd /d %~dp0
python win10_red_monitor_sleep.py
if errorlevel 1 (
  echo.
  echo 程序异常退出，请检查上方报错信息。
  pause
)
