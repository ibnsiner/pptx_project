@echo off
setlocal EnableExtensions
REM Use ASCII only: UTF-8 Korean breaks cmd.exe on some systems.

pushd "%~dp0parser-api" || (
  echo ERROR: cannot open folder parser-api
  pause
  exit /b 1
)

if not exist ".venv\Scripts\activate.bat" (
  echo ERROR: missing parser-api\.venv
  echo Run once:
  echo   python -m venv .venv
  echo   .venv\Scripts\activate.bat
  echo   pip install -r requirements.txt
  popd
  pause
  exit /b 1
)

call ".venv\Scripts\activate.bat"

echo.
echo ---- parser-api folder (must end with \parser-api^) ----
echo %CD%
echo ---- expected: ...\PPTX_Parsing\parser-api
echo.

python -c "import pathlib; p=pathlib.Path('app/main.py').resolve(); print('app.main path:', p); assert p.name=='main.py', p"

python -c "from app.main import PARSER_API_BUILD; print('PARSER_API_BUILD:', PARSER_API_BUILD)"

echo.
echo Default port is 8010 so it does not clash with another app on 8000.
echo frontend/.env.local must have PARSER_API_URL=http://127.0.0.1:8010
echo Test: http://127.0.0.1:8010/health  and  http://127.0.0.1:8010/pptx-parser/build
echo.
echo parser-api: http://127.0.0.1:8010  ^(Ctrl+C to stop^)
python -m uvicorn app.main:app --reload --host 127.0.0.1 --port 8010
popd
pause
