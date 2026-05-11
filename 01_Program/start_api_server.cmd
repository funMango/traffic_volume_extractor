@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..") do set "PROJECT_ROOT=%%~fI"

set "PYTHON_EXE=C:\Users\An JaeYeol\AppData\Local\Programs\Python\Python312\python.exe"
set "PYTHON_CMD="
set "API_MAIN=%SCRIPT_DIR%traffic_api\main.py"
set "LOG_FILE=%PROJECT_ROOT%\02_Result\api_autostart.log"

if not exist "%PROJECT_ROOT%\02_Result" (
    mkdir "%PROJECT_ROOT%\02_Result" >nul 2>&1
)

echo [%date% %time%] [INFO] Starting traffic_api server>>"%LOG_FILE%"

if exist "%PYTHON_EXE%" (
    set "PYTHON_CMD="%PYTHON_EXE%""
)

if not defined PYTHON_CMD (
    py -3 --version >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=py -3"
)

if not defined PYTHON_CMD (
    python --version >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=python"
)

if not defined PYTHON_CMD (
    echo [%date% %time%] [ERROR] Python not found. Tried: %PYTHON_EXE%, py -3, python>>"%LOG_FILE%"
    exit /b 9009
)

if not exist "%API_MAIN%" (
    echo [%date% %time%] [ERROR] traffic_api main not found: %API_MAIN%>>"%LOG_FILE%"
    exit /b 2
)

pushd "%SCRIPT_DIR%" >nul 2>&1
%PYTHON_CMD% -m traffic_api.main >>"%LOG_FILE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"
popd >nul 2>&1

echo [%date% %time%] [INFO] traffic_api server exited with code %EXIT_CODE%>>"%LOG_FILE%"
exit /b %EXIT_CODE%
