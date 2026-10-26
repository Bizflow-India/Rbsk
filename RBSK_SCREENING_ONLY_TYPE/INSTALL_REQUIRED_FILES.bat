@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo Checking and installing Python and required packages...

call "%~dp0ONE_CLICK_INSTALL_AND_RUN.bat"
