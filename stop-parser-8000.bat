@echo off
setlocal EnableExtensions
REM Free port 8000. From PowerShell: .\stop-parser-8000.bat

cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0stop-parser-8000.ps1"
echo.
pause
