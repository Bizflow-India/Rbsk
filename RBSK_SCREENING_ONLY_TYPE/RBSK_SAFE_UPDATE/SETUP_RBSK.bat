@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

title RBSK Screening - Automatic Setup
color 0B

echo.
echo  ============================================================
echo     RBSK Screening Health - Automatic Setup
echo  ============================================================
echo.
echo   This will install everything needed to run the RBSK
echo   Screening application on this PC.
echo.
echo  ============================================================
echo.

set SCRIPT_DIR=%~dp0
set PY_CMD=

REM ── Step 1: Check if Python is installed ──
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

REM ── Python missing: Install automatically ──
echo.
echo   Python is NOT installed on this PC. Installing automatically...
echo   Please wait...
echo.

set INSTALLER_EXE=%SCRIPT_DIR%python_installer\python-3.12.8-amd64.exe

if not exist "%INSTALLER_EXE%" (
    echo   Downloading official Python installer...
    if not exist "%SCRIPT_DIR%python_installer" mkdir "%SCRIPT_DIR%python_installer"
    powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe' -OutFile '%INSTALLER_EXE%'"
)

if exist "%INSTALLER_EXE%" (
    echo   Installing Python silently...
    "%INSTALLER_EXE%" /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1 SimpleInstall=1
    timeout /t 5 /nobreak >nul
) else (
    color 0C
    echo   ERROR: Could not download or find Python installer.
    pause
    exit /b 1
)

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

:PYTHON_FOUND
echo   Found Python: %PY_CMD%
for /f "tokens=*" %%v in ('"%PY_CMD%" --version 2^>^&1') do set PYVER=%%v
echo   Version: %PYVER%

REM ── Step 2: Upgrade pip ──
echo.
echo [2/4] Updating pip...
"%PY_CMD%" -m pip install --upgrade pip --quiet >nul 2>nul

REM ── Step 3: Install required packages ──
echo.
echo [3/4] Installing required packages (selenium, cryptography, openpyxl, requests, urllib3)...

if exist "%SCRIPT_DIR%wheels" (
    echo   Installing packages from local offline wheels directory...
    "%PY_CMD%" -m pip install --no-index --find-links="%SCRIPT_DIR%wheels" selenium cryptography openpyxl requests urllib3 --quiet >nul 2>nul
)

"%PY_CMD%" -c "import selenium, cryptography, openpyxl, requests, urllib3" >nul 2>nul
if errorlevel 1 (
    echo   Installing remaining packages online...
    "%PY_CMD%" -m pip install cryptography openpyxl selenium requests urllib3 --quiet
)

"%PY_CMD%" -c "import selenium, cryptography, openpyxl, requests, urllib3" >nul 2>nul
if errorlevel 1 (
    color 0C
    echo   WARNING: Some packages failed to install.
    pause
    exit /b 1
) else (
    echo   [OK] All packages installed successfully.
)

REM ── Step 4: Create Desktop Shortcut ──
echo.
echo [4/4] Creating desktop shortcut...
set SHORTCUT_NAME=RBSK Screening.lnk

if exist "%SCRIPT_DIR%create_shortcut.ps1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%create_shortcut.ps1" >nul 2>nul
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $d=[Environment]::GetFolderPath('Desktop'); $sc=$ws.CreateShortcut([System.IO.Path]::Combine($d,'RBSK Screening.lnk')); $sc.TargetPath='cmd.exe'; $sc.Arguments='/c \"\"%SCRIPT_DIR%OPEN_RBSK_SCREENING_APP.bat\"\"'; $sc.WorkingDirectory='%SCRIPT_DIR%'; $sc.IconLocation='shell32.dll,14'; $sc.Description='RBSK Screening Health Application'; $sc.Save()" >nul 2>nul
)

if not errorlevel 1 (
    echo   [OK] Desktop shortcut created.
) else (
    echo   [SKIP] Could not create desktop shortcut.
)

echo.
color 0A
echo  ============================================================
echo     SETUP COMPLETE! All packages installed successfully.
echo  ============================================================
echo.
echo   The RBSK Screening app will now launch automatically.
echo.
echo  ============================================================
echo.

start "" "%SCRIPT_DIR%OPEN_RBSK_SCREENING_APP.bat"

timeout /t 3 /nobreak >nul
exit /b 0
