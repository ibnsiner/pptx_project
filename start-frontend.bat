@echo off
setlocal EnableExtensions
REM Use ASCII only: UTF-8 Korean breaks cmd.exe on some systems.

pushd "%~dp0frontend" || (
  echo ERROR: cannot open folder frontend
  pause
  exit /b 1
)

if not exist "node_modules\" (
  echo Running npm install...
  call npm install
  if errorlevel 1 (
    echo npm install failed
    popd
    pause
    exit /b 1
  )
)

echo frontend: Next.js dev server ^(Ctrl+C to stop^)
call npm run dev
popd
pause
