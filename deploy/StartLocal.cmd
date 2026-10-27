@echo off
powershell.exe -NoProfile -File "%~dp0local.ps1" -Action start
if errorlevel 1 (
  echo See guide.html for help. No execution policy change is required.
  pause
  exit /b 1
)
if not defined JAKAL_LOCAL_PORT set "JAKAL_LOCAL_PORT=18988"
start "" "http://127.0.0.1:%JAKAL_LOCAL_PORT%/"
