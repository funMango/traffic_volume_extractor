@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..") do set "PROJECT_ROOT=%%~fI"

set "VENV_PYTHON=%PROJECT_ROOT%\.venv\Scripts\python.exe"
set "USER_PYTHON=C:\Users\An JaeYeol\AppData\Local\Programs\Python\Python312\python.exe"
set "PYTHON_CMD="
set "PYTHON_SOURCE="
set "API_MAIN=%SCRIPT_DIR%traffic_api\main.py"
set "LOG_FILE=%PROJECT_ROOT%\02_Result\api_autostart.log"
set "SERVER_LOG_FILE=%PROJECT_ROOT%\02_Result\server.log"
set "PORT_LISTENING="

if not exist "%PROJECT_ROOT%\02_Result" (
    mkdir "%PROJECT_ROOT%\02_Result" >nul 2>&1
)

echo [%date% %time%] [INFO] Starting traffic_api server>>"%LOG_FILE%"
echo [%date% %time%] [INFO] Project root: %PROJECT_ROOT%>>"%LOG_FILE%"
echo [%date% %time%] [INFO] Working directory: %SCRIPT_DIR%>>"%LOG_FILE%"

for /f "tokens=*" %%L in ('netstat -ano ^| findstr /r /c:":8000 .*LISTENING"') do (
    set "PORT_LISTENING=1"
    echo [%date% %time%] [INFO] Port 8000 already listening: %%L>>"%LOG_FILE%"
)

if defined PORT_LISTENING (
    echo [%date% %time%] [INFO] traffic_api server is already running; skipping duplicate start>>"%LOG_FILE%"
    exit /b 0
)

if exist "%VENV_PYTHON%" (
    call :CHECK_PYTHON "%VENV_PYTHON%"
    if not errorlevel 1 (
        set "PYTHON_CMD="%VENV_PYTHON%""
        set "PYTHON_SOURCE=project venv"
        goto PYTHON_SELECTED
    ) else (
        echo [%date% %time%] [WARN] Python candidate missing required modules: "%VENV_PYTHON%">>"%LOG_FILE%"
    )
)

if not defined PYTHON_CMD if exist "%USER_PYTHON%" (
    call :CHECK_PYTHON "%USER_PYTHON%"
    if not errorlevel 1 (
        set "PYTHON_CMD="%USER_PYTHON%""
        set "PYTHON_SOURCE=user python"
        goto PYTHON_SELECTED
    ) else (
        echo [%date% %time%] [WARN] Python candidate missing required modules: "%USER_PYTHON%">>"%LOG_FILE%"
    )
)

if not defined PYTHON_CMD (
    py -3 -c "import fastapi, uvicorn" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_CMD=py -3"
        set "PYTHON_SOURCE=py launcher"
        goto PYTHON_SELECTED
    ) else (
        echo [%date% %time%] [WARN] Python candidate missing required modules or unavailable: py -3>>"%LOG_FILE%"
    )
)

if not defined PYTHON_CMD (
    python -c "import fastapi, uvicorn" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_CMD=python"
        set "PYTHON_SOURCE=PATH python"
        goto PYTHON_SELECTED
    ) else (
        echo [%date% %time%] [WARN] Python candidate missing required modules or unavailable: python>>"%LOG_FILE%"
    )
)

:PYTHON_SELECTED

if not defined PYTHON_CMD (
    echo [%date% %time%] [ERROR] Usable Python not found. Tried: %VENV_PYTHON%, %USER_PYTHON%, py -3, python>>"%LOG_FILE%"
    exit /b 9009
)

if not exist "%API_MAIN%" (
    echo [%date% %time%] [ERROR] traffic_api main not found: %API_MAIN%>>"%LOG_FILE%"
    exit /b 2
)

echo [%date% %time%] [INFO] Selected Python: %PYTHON_CMD% (%PYTHON_SOURCE%)>>"%LOG_FILE%"

pushd "%SCRIPT_DIR%" >nul 2>&1
%PYTHON_CMD% -m traffic_api.main >>"%SERVER_LOG_FILE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"
popd >nul 2>&1

echo [%date% %time%] [INFO] traffic_api server exited with code %EXIT_CODE%>>"%LOG_FILE%"
exit /b %EXIT_CODE%

:CHECK_PYTHON
%1 -c "import fastapi, uvicorn" >nul 2>&1
exit /b %ERRORLEVEL%
