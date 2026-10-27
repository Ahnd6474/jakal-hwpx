@echo off
if not exist "%~dp0Jakal.exe" (
  echo Extract the entire ZIP before starting Jakal.
  pause
  exit /b 1
)
start "" "%~dp0Jakal.exe"
