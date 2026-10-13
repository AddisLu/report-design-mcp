@echo off
setlocal
rem Offline install: no PyPI access needed. Usage: install_offline.bat [path\to\python.exe]
cd /d "%~dp0"

set "PY=%~1"
if "%PY%"=="" set "PY=python"

"%PY%" -c "import sys,struct; v=sys.version_info; b=struct.calcsize('P')*8; print('Python %%d.%%d.%%d %%d-bit' %% (v[0],v[1],v[2],b)); sys.exit(0 if (v[:2]==(3,8) and b==64) else 1)"
if errorlevel 1 (
  echo.
  echo [ERROR] Need 64-bit Python 3.8. Pass its path, e.g.:
  echo     install_offline.bat C:\Python38\python.exe
  exit /b 1
)

if not exist .venv\Scripts\python.exe (
  echo Creating .venv ...
  "%PY%" -m venv .venv
  if errorlevel 1 exit /b 1
)

echo Installing packages from wheelhouse (offline) ...
.venv\Scripts\python -m pip install --no-index --find-links wheelhouse -r requirements.txt
if errorlevel 1 (
  echo [ERROR] pip install failed. Send the output above.
  exit /b 1
)

echo Running smoke test ...
.venv\Scripts\python tests\smoke_test.py
exit /b %errorlevel%
