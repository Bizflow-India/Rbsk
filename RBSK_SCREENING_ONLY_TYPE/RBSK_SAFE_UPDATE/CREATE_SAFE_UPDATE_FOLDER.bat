@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0create_safe_update_folder.ps1"
if errorlevel 1 (
    echo.
    echo SAFE UPDATE FOLDER CREATION FAILED.
    pause
    exit /b 1
)
echo.
echo Press any key to open the safe update folder.
pause >nul
explorer "%~dp0RBSK_SAFE_UPDATE"

