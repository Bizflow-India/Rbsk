@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

title RBSK Screening - 1-Click Auto Setup & Launcher
color 0B

echo.
echo  ============================================================
echo     RBSK Screening Health - 1-Click Installer & Launcher
echo  ============================================================
echo.
echo   This setup will automatically install Python, Selenium, and
echo   all required components, create a desktop shortcut, and
echo   launch the RBSK Screening app!
echo.
echo  ============================================================
echo.

set SCRIPT_DIR=%~dp0
set PY_CMD=

REM ── Step 1: Detect existing Python installation ──
echo [1/4] Checking Python installation...

where python >nul 2>nul
if not errorlevel 1 (
    set "PY_CMD=python"
    goto :PYTHON_FOUND
)

where pythonw >nul 2>nul
if not errorlevel 1 (
    set "PY_CMD=pythonw"
    goto :PYTHON_FOUND
)

REM Check common Python install locations
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python312\python.exe"
    set "PATH=%LocalAppData%\Programs\Python\Python312;%LocalAppData%\Programs\Python\Python312\Scripts;%PATH%"
    goto :PYTHON_FOUND
)
if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python313\python.exe"
    set "PATH=%LocalAppData%\Programs\Python\Python313;%LocalAppData%\Programs\Python\Python313\Scripts;%PATH%"
    goto :PYTHON_FOUND
)
if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python311\python.exe"
    set "PATH=%LocalAppData%\Programs\Python\Python311;%LocalAppData%\Programs\Python\Python311\Scripts;%PATH%"
    goto :PYTHON_FOUND
)
if exist "C:\Program Files\Python312\python.exe" (
    set "PY_CMD=C:\Program Files\Python312\python.exe"
    set "PATH=C:\Program Files\Python312;C:\Program Files\Python312\Scripts;%PATH%"
    goto :PYTHON_FOUND
)
if exist "C:\Python312\python.exe" (
    set "PY_CMD=C:\Python312\python.exe"
    set "PATH=C:\Python312;C:\Python312\Scripts;%PATH%"
    goto :PYTHON_FOUND
)

REM ── Python not found: Install Python automatically ──
echo.
echo   Python is NOT installed. Installing Python automatically...
echo   Please wait, this will take about 30 seconds...
echo.

set INSTALLER_EXE=%SCRIPT_DIR%python_installer\python-3.12.8-amd64.exe

if not exist "%INSTALLER_EXE%" (
    echo   [+] Python installer file not found locally. Downloading official Python installer...
    if not exist "%SCRIPT_DIR%python_installer" mkdir "%SCRIPT_DIR%python_installer"
    powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe' -OutFile '%INSTALLER_EXE%'"
)

if exist "%INSTALLER_EXE%" (
    echo   [+] Running silent Python installer...
    "%INSTALLER_EXE%" /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1 SimpleInstall=1
    timeout /t 5 /nobreak >nul
) else (
    color 0C
    echo   [ERROR] Could not find or download Python installer.
    echo   Please make sure you have internet connection or place python-3.12.8-amd64.exe in python_installer folder.
    pause
    exit /b 1
)

REM Refresh path check after install
set "PATH=%LocalAppData%\Programs\Python\Python312;%LocalAppData%\Programs\Python\Python312\Scripts;C:\Program Files\Python312;C:\Program Files\Python312\Scripts;%PATH%"
where python >nul 2>nul
if not errorlevel 1 (
    set "PY_CMD=python"
    goto :PYTHON_FOUND
)
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python312\python.exe"
    goto :PYTHON_FOUND
)
if exist "C:\Program Files\Python312\python.exe" (
    set "PY_CMD=C:\Program Files\Python312\python.exe"
    goto :PYTHON_FOUND
)

:PYTHON_FOUND
echo   [OK] Python located: %PY_CMD%
for /f "tokens=*" %%v in ('"%PY_CMD%" --version 2^>^&1') do set PYVER=%%v
echo        Version: %PYVER%

REM ── Step 2: Install required libraries (Selenium, cryptography, openpyxl, etc.) ──
echo.
echo [2/4] Installing required packages (Selenium, cryptography, openpyxl, requests, urllib3)...

if exist "%SCRIPT_DIR%wheels" (
    echo   [+] Installing from local offline wheel packages...
    "%PY_CMD%" -m pip install --no-index --find-links="%SCRIPT_DIR%wheels" selenium cryptography openpyxl requests urllib3 --quiet >nul 2>nul
)

REM Verify selenium and packages
"%PY_CMD%" -c "import selenium, cryptography, openpyxl, requests, urllib3" >nul 2>nul
if errorlevel 1 (
    echo   [+] Wheel install incomplete. Installing via pip network...
    "%PY_CMD%" -m pip install selenium cryptography openpyxl requests urllib3 --quiet
)

"%PY_CMD%" -c "import selenium, cryptography, openpyxl, requests, urllib3" >nul 2>nul
if errorlevel 1 (
    color 0C
    echo.
    echo   [ERROR] Failed to install required Python packages.
    echo   Please verify your internet connection or wheels folder.
    pause
    exit /b 1
) else (
    echo   [OK] Selenium and all required packages are ready!
)

REM ── Step 3: Create Desktop Shortcut ──
echo.
echo [3/4] Creating Desktop Shortcut...
set SHORTCUT_NAME=RBSK Screening.lnk

if exist "%SCRIPT_DIR%create_shortcut.ps1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%create_shortcut.ps1" >nul 2>nul
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $d=[Environment]::GetFolderPath('Desktop'); $sc=$ws.CreateShortcut([System.IO.Path]::Combine($d,'RBSK Screening.lnk')); $sc.TargetPath='cmd.exe'; $sc.Arguments='/c \"\"%SCRIPT_DIR%OPEN_RBSK_SCREENING_APP.bat\"\"'; $sc.WorkingDirectory='%SCRIPT_DIR%'; $sc.IconLocation='shell32.dll,14'; $sc.Description='RBSK Screening Health Application'; $sc.Save()" >nul 2>nul
)

if not errorlevel 1 (
    echo   [OK] Desktop shortcut "RBSK Screening" created successfully.
) else (
    echo   [SKIP] Desktop shortcut creation skipped.
)

REM ── Step 4: Launch Application ──
echo.
echo [4/4] Launching RBSK Screening Application...
color 0A
echo.
echo  ============================================================
echo     SETUP & INSTALLATION COMPLETE! Launching app...
echo  ============================================================
echo.

start "" "%SCRIPT_DIR%OPEN_RBSK_SCREENING_APP.bat"

timeout /t 3 /nobreak >nul
exit /b 0
