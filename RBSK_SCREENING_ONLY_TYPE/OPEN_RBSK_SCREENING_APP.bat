@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set SCRIPT_DIR=%~dp0
set PY_CMD=

where pythonw >nul 2>nul
if not errorlevel 1 (
    set "PY_CMD=pythonw"
    goto :RUN_APP
)

where python >nul 2>nul
if not errorlevel 1 (
    set "PY_CMD=python"
    goto :RUN_APP
)

if exist "%LocalAppData%\Programs\Python\Python312\pythonw.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python312\pythonw.exe"
    goto :RUN_APP
)
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python312\python.exe"
    goto :RUN_APP
)
if exist "%LocalAppData%\Programs\Python\Python313\pythonw.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python313\pythonw.exe"
    goto :RUN_APP
)
if exist "%LocalAppData%\Programs\Python\Python311\pythonw.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python311\pythonw.exe"
    goto :RUN_APP
)
if exist "C:\Program Files\Python312\pythonw.exe" (
    set "PY_CMD=C:\Program Files\Python312\pythonw.exe"
    goto :RUN_APP
)
if exist "C:\Python312\pythonw.exe" (
    set "PY_CMD=C:\Python312\pythonw.exe"
    goto :RUN_APP
)

echo Python is not detected. Starting 1-Click setup...
call "%SCRIPT_DIR%ONE_CLICK_INSTALL_AND_RUN.bat"
exit /b 0

:RUN_APP
start "" "%PY_CMD%" "%SCRIPT_DIR%RBSK_Easy_Launcher.pyw"
exit /b 0
