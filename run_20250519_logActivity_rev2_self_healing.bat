@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM === Work from the folder that contains this BAT ===
set "ROOT_DIR=%~dp0"
pushd "%ROOT_DIR%"

REM === Config ===
REM This BAT assumes the .bat, .py, and "Log Activity.xlsx" live together.
set "ENV_NAME=ENV_SAP_Log_Activity"
set "ENV_DIR=%ROOT_DIR%%ENV_NAME%"
set "ENV_ACT=%ENV_DIR%\Scripts\activate.bat"
set "ENV_PY=%ENV_DIR%\Scripts\python.exe"
set "SCRIPT=%ROOT_DIR%20250519_logActivity_rev2.py"

REM --- Sanity: script present? ---
if not exist "%SCRIPT%" (
  echo [ERROR] Could not find script: "%SCRIPT%"
  echo Ensure this BAT sits next to 20250519_logActivity_rev2.py
  popd & endlocal & exit /b 1
)

REM --- Resolve a Python 3.11+ command (launcher, local python.exe, or PATH) ---
set "PY_CMD="
set "LOCAL_PY=%ROOT_DIR%python.exe"

if exist "%LOCAL_PY%" (
  set "PY_CMD=%LOCAL_PY%"
) else (
  py -3.11 --version >nul 2>&1 && set "PY_CMD=py -3.11"
)

if not defined PY_CMD (
  where python >nul 2>&1 && set "PY_CMD=python"
)

if not defined PY_CMD (
  echo [ERROR] Python not found. Install Python 3.11+ or place python.exe next to this BAT.
  popd & endlocal & exit /b 1
)

REM --- Confirm Python 3.11+ ---
%PY_CMD% -c "import sys; raise SystemExit(0 if sys.version_info[:2] >= (3,11) else 1)" >nul 2>&1 || (
  echo [ERROR] Python 3.11+ required. Found a lower version.
  echo Install Python 3.11.x or adjust PATH / local python.exe.
  popd & endlocal & exit /b 1
)

REM --- Create venv if missing; otherwise repair if it's broken ---
if not exist "%ENV_PY%" (
  echo [INFO] Creating venv "%ENV_NAME%" with Python 3.11+ ...
  %PY_CMD% -m venv "%ENV_DIR%"
  if errorlevel 1 (
    echo [ERROR] venv creation failed.
    popd & endlocal & exit /b 1
  )
) else (
  REM Try to run the venv's python; if that fails, repair pointers
  "%ENV_PY%" --version >nul 2>&1 || (
    echo [WARN] venv Python not responding; repairing with --upgrade ...
    %PY_CMD% -m venv --upgrade "%ENV_DIR%"
  )
)

REM --- Final check on venv python ---
"%ENV_PY%" --version >nul 2>&1 || (
  echo [ERROR] venv Python still not usable after repair.
  popd & endlocal & exit /b 1
)

REM === Dependencies: install only if missing (idempotent) ===
set "NEED_WDM="
findstr /I /C:"webdriver_manager.chrome" "%SCRIPT%" >nul 2>&1 && set "NEED_WDM=1"

REM Ensure core packages are present
"%ENV_PY%" -c "import selenium, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo [INFO] Installing core packages into the venv...
  "%ENV_PY%" -m pip install --upgrade pip
  "%ENV_PY%" -m pip install "selenium==4.21.0" "openpyxl==3.1.2"
)

if defined NEED_WDM (
  "%ENV_PY%" -c "import webdriver_manager" >nul 2>&1 || (
    echo [INFO] Installing webdriver-manager because the script imports it...
    "%ENV_PY%" -m pip install "webdriver-manager==4.0.2"
  )
  REM Optional: clear old/bad driver cache that can cause WinError 193
  if exist "%USERPROFILE%\.wdm\drivers\chromedriver" (
    echo [INFO] Clearing webdriver-manager driver cache...
    rmdir /s /q "%USERPROFILE%\.wdm\drivers\chromedriver"
  )
)

REM --- Activate venv and run the script ---
call "%ENV_ACT%" || (
  echo [ERROR] Failed to activate virtual environment.
  popd & endlocal & exit /b 1
)

echo.
echo [RUN] "%SCRIPT%"
"%ENV_PY%" "%SCRIPT%"
set "RC=%ERRORLEVEL%"
echo.

if not "%RC%"=="0" (
  echo [EXIT %RC%] Script finished with errors.
) else (
  echo [OK] Script completed successfully.
)

pause
popd
endlocal
exit /b %RC%
