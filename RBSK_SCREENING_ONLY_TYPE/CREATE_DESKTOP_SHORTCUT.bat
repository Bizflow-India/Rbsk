@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

title RBSK Screening - Create Desktop Shortcut
color 0A

echo.
echo  ============================================================
echo     Creating RBSK Screening Desktop Shortcut
echo  ============================================================
echo.

if exist "%~dp0create_shortcut.ps1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0create_shortcut.ps1"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $d=[Environment]::GetFolderPath('Desktop'); $sc=$ws.CreateShortcut([System.IO.Path]::Combine($d,'RBSK Screening.lnk')); $sc.TargetPath='cmd.exe'; $sc.Arguments='/c \"\"%~dp0OPEN_RBSK_SCREENING_APP.bat\"\"'; $sc.WorkingDirectory='%~dp0'; $sc.IconLocation='shell32.dll,14'; $sc.Description='RBSK Screening Health Application'; $sc.Save(); Write-Host '  [OK] Desktop shortcut created successfully!'"
)

echo.
echo  ============================================================
echo.
timeout /t 5
