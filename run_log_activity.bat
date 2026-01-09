@echo off
setlocal

REM Go to the folder where this .bat lives
pushd "%~dp0"

REM --- Paths ---
set "VENV_DIR=%~dp0SAP_Log_Activity_250808_a"
set "VENV_ACT=%VENV_DIR%\Scripts\activate.bat"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "SCRIPT=%~dp020250519_logActivity_rev2.py"

REM --- Sanity checks ---
if not exist "%VENV_ACT%" (
  echo [ERROR] Missing: "%VENV_ACT%"
  popd & endlocal & exit /b 1
)
if not exist "%VENV_PY%" (
  echo [ERROR] Missing: "%VENV_PY%"
  echo Try: py -3.11 -m venv --upgrade "%VENV_DIR%"
  popd & endlocal & exit /b 1
)
if not exist "%SCRIPT%" (
  echo [ERROR] Missing script: "%SCRIPT%"
  popd & endlocal & exit /b 1
)

REM --- Activate the venv and run the script ---
call "%VENV_ACT%" || (
  echo [ERROR] Failed to activate virtual environment.
  popd & endlocal & exit /b 1
)

"%VENV_PY%" "%SCRIPT%"
set "RC=%ERRORLEVEL%"

popd
endlocal
exit /b %RC%
