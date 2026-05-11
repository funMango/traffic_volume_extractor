@echo off
setlocal

set "PYTHON_WRAPPER=%~dp0python.ps1"

if not exist "%PYTHON_WRAPPER%" (
  echo Python wrapper not found: %PYTHON_WRAPPER% 1>&2
  exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PYTHON_WRAPPER%" %*
exit /b %ERRORLEVEL%
