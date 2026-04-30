@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..") do set "PROJECT_ROOT=%%~fI"

set "PYTHON_EXE=C:\Users\An JaeYeol\AppData\Local\Programs\Python\Python312\python.exe"
set "API_SCRIPT=%SCRIPT_DIR%api_server.py"
set "LOG_FILE=%PROJECT_ROOT%\02_Result\api_autostart.log"

if not exist "%PROJECT_ROOT%\02_Result" (
    mkdir "%PROJECT_ROOT%\02_Result" >nul 2>&1
)

echo [%date% %time%] [INFO] Starting API server>>"%LOG_FILE%"

if not exist "%PYTHON_EXE%" (
    echo [%date% %time%] [ERROR] Python not found: %PYTHON_EXE%>>"%LOG_FILE%"
    exit /b 9009
)

if not exist "%API_SCRIPT%" (
    echo [%date% %time%] [ERROR] api_server.py not found: %API_SCRIPT%>>"%LOG_FILE%"
    exit /b 2
)

"%PYTHON_EXE%" "%API_SCRIPT%" >>"%LOG_FILE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"

echo [%date% %time%] [INFO] API server exited with code %EXIT_CODE%>>"%LOG_FILE%"
exit /b %EXIT_CODE%
