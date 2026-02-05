@echo off
setlocal

REM Go to the folder where this .bat lives
pushd "%~dp0"

REM --- Paths ---
set "VENV_DIR=%~dp0logActivity_venv"
set "VENV_ACT=%VENV_DIR%\Scripts\activate.bat"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "SCRIPT=%~dp020260108_logActivity.py"

REM --- Sanity checks ---
if not exist "%VENV_ACT%" (
  echo [INFO] Creating virtual environment at: "%VENV_DIR%"
  py -3.11 -m venv "%VENV_DIR%" || (
    echo [ERROR] Failed to create virtual environment.
    popd & endlocal & exit /b 1
  )
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
