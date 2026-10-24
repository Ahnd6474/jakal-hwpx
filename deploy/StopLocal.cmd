@echo off
powershell.exe -NoProfile -File "%~dp0local.ps1" -Action stop
if errorlevel 1 (
  pause
  exit /b 1
)
