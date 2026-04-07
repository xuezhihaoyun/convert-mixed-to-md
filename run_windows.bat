@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "PYTHONUTF8=1"
set "PYTHONDONTWRITEBYTECODE=1"

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%convert_mixed_to_md.py"
set "REQUIREMENTS_PATH=%SCRIPT_DIR%requirements.txt"
set "PATH_PARSER=%SCRIPT_DIR%parse_windows_input_paths.py"
set "VENV_DIR=%SCRIPT_DIR%.venv"
set "VENV_PYTHON=%VENV_DIR%\Scripts\python.exe"

if not exist "%SCRIPT_PATH%" (
  echo [FAIL] Script not found: %SCRIPT_PATH%
  pause
  exit /b 1
)

set "PYTHON_LAUNCHER="
where py >nul 2>&1
if not errorlevel 1 set "PYTHON_LAUNCHER=py -3"
if not defined PYTHON_LAUNCHER (
  where python >nul 2>&1
  if not errorlevel 1 set "PYTHON_LAUNCHER=python"
)
if not defined PYTHON_LAUNCHER (
  echo [FAIL] Python 3 is not found in PATH.
  pause
  exit /b 1
)

call :ensure_runtime
if errorlevel 1 (
  pause
  exit /b 1
)

where pandoc >nul 2>&1
if errorlevel 1 echo [WARN] pandoc not found. doc/docx/epub conversion may fail.
where pdftotext >nul 2>&1
if errorlevel 1 echo [WARN] pdftotext not found. PDF extraction will fall back to python extractors.

set /a SUCCESS_COUNT=0
set /a FAIL_COUNT=0
set "WARNED_TOKEN=0"

echo convert_mixed_to_md (Windows)
echo Supported: doc / docx / pdf / epub / wps / wpt / hwp
echo You can drag file/folder to this .bat, or input path manually.

if not "%~1"=="" (
  :args_loop
  if "%~1"=="" goto interactive_loop
  call :dispatch_input "%~1"
  shift
  goto args_loop
)

:interactive_loop
echo.
set /p TARGET_PATH_RAW=Input file/folder path (Enter to finish): 
if "%TARGET_PATH_RAW%"=="" goto done
call :dispatch_input "%TARGET_PATH_RAW%"
echo Current: success !SUCCESS_COUNT!, fail !FAIL_COUNT!
goto interactive_loop

:done
echo.
echo Finished: success !SUCCESS_COUNT!, fail !FAIL_COUNT!
pause
if !FAIL_COUNT! gtr 0 exit /b 1
exit /b 0

:dispatch_input
set "RAW_INPUT=%~1"
if "%RAW_INPUT%"=="" exit /b 0
set "PARSED_ANY=0"
if exist "%PATH_PARSER%" (
  for /f "usebackq delims=" %%P in (`"%VENV_PYTHON%" "%PATH_PARSER%" "%RAW_INPUT%"`) do (
    set "PARSED_ANY=1"
    call :process_one "%%P"
  )
)
if "%PARSED_ANY%"=="0" (
  call :process_one "%RAW_INPUT%"
)
exit /b 0

:process_one
set "TARGET_PATH=%~1"
if "%TARGET_PATH%"=="" exit /b 0
echo.
echo Processing: %TARGET_PATH%
if not defined MINERU_TOKEN if "!WARNED_TOKEN!"=="0" (
  echo [INFO] MINERU_TOKEN is not set. Scanned PDF OCR may fail.
  set "WARNED_TOKEN=1"
)
"%VENV_PYTHON%" "%SCRIPT_PATH%" "%TARGET_PATH%"
if errorlevel 1 (
  set /a FAIL_COUNT+=1
  echo [FAIL] %TARGET_PATH%
) else (
  set /a SUCCESS_COUNT+=1
  echo [OK] %TARGET_PATH%
)
exit /b 0

:ensure_runtime
if not exist "%VENV_PYTHON%" (
  echo [INIT] Creating local venv...
  %PYTHON_LAUNCHER% -m venv "%VENV_DIR%"
  if errorlevel 1 (
    echo [FAIL] Failed to create venv.
    exit /b 1
  )
)

"%VENV_PYTHON%" -c "import requests, pdfplumber, pypdf" >nul 2>&1
if errorlevel 1 (
  echo [INIT] Installing dependencies...
  "%VENV_PYTHON%" -m pip install --upgrade pip >nul 2>&1
  "%VENV_PYTHON%" -m pip install -r "%REQUIREMENTS_PATH%"
  if errorlevel 1 (
    echo [FAIL] Dependency install failed.
    echo Run manually:
    echo "%VENV_PYTHON%" -m pip install -r "%REQUIREMENTS_PATH%"
    exit /b 1
  )
)

"%VENV_PYTHON%" -c "import requests, pdfplumber, pypdf" >nul 2>&1
if errorlevel 1 (
  echo [FAIL] Python dependencies are still unavailable.
  exit /b 1
)
exit /b 0
