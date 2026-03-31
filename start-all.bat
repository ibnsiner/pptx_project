@echo off
setlocal EnableExtensions
REM Use ASCII only: UTF-8 Korean breaks cmd.exe on some systems.

cd /d "%~dp0"
echo Starting backend and frontend in new windows...
echo Backend: http://127.0.0.1:8010
echo Frontend: URL printed in the other window ^(often http://localhost:3000^)
echo.

start "PPTX Parser API" cmd /k "%~dp0start-backend.bat"
timeout /t 2 /nobreak >nul
start "PPTX Frontend" cmd /k "%~dp0start-frontend.bat"

echo Two windows should be open. You can close this one.
pause
