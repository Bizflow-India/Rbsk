@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python is not installed or not added to PATH.
  echo.
  echo Please install Python from:
  echo https://www.python.org/downloads/
  echo.
  echo IMPORTANT: During install, tick "Add python.exe to PATH".
  pause
  exit /b 1
)

echo Checking required Python files...
python -c "import cryptography" >nul 2>nul
if errorlevel 1 (
  echo Installing cryptography...
  python -m pip install cryptography
  if errorlevel 1 (
    echo Failed to install cryptography.
    pause
    exit /b 1
  )
)

echo.
echo Required files are ready.
echo You can now double-click OPEN_KEY_GENERATOR.bat
pause
