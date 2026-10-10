@echo off
setlocal
cd /d "%~dp0"
set "PACKING_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%~dp0.venv\Scripts\python.exe" set "PACKING_PYTHON=%~dp0.venv\Scripts\python.exe"
if exist "%PACKING_PYTHON%" (
  "%PACKING_PYTHON%" "%~dp0PackingListApp\src\app.py" %*
) else (
  py -3.12 "%~dp0PackingListApp\src\app.py" %*
)
if errorlevel 1 (
  echo.
  echo Packing List Manager could not start. The message above explains why.
  echo This source version needs Python 3.12 and the packages in PackingListApp\requirements.txt.
  pause
)
endlocal
