@echo off
setlocal
cd /d "%~dp0"

where pythonw >nul 2>nul
if errorlevel 1 (
  echo Python is not installed or not added to PATH.
  echo Install Python first, then run this file again.
  pause
  exit /b 1
)

start "" pythonw "%~dp0RBSK_License_Key_Generator.pyw"
