
@echo off
setlocal
title Gmail Labeler - Build

cd /d "%~dp0"

echo.
echo ==========================================
echo          GMAIL LABELER - BUILD
echo ==========================================
echo.
echo Project directory:
echo %CD%
echo.

if not exist "src\Setup-GmailLabeler.ps1" (
    echo ERROR: PowerShell setup script not found!
    goto ERROR
)

if not exist "assets\icon.ico" (
    echo ERROR: Application icon not found!
    goto ERROR
)

REM Pass the local directory to PowerShell
set "PROJECT_DIR=%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\Build-GmailLabeler.ps1"

if errorlevel 1 goto ERROR

echo.
echo ==========================================
echo           BUILD COMPLETED!
echo ==========================================
echo.
pause
exit /b 0

:ERROR
echo.
echo ==========================================
echo             BUILD FAILED!
echo ==========================================
echo.
pause
exit /b 1
