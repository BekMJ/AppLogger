@echo off
setlocal

REM Path to the app you want to launch after logging
set "APP_PATH=C:\Program Files\FlowVision\FlowVision.exe"
set "APP_ARGS="

REM Derive script directory to locate AppLogger.ps1
set "SCRIPT_DIR=%~dp0"
set "LOGGER_PS=%SCRIPT_DIR%AppLogger.ps1"

if not exist "%LOGGER_PS%" (
  echo AppLogger.ps1 not found in %SCRIPT_DIR%
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%LOGGER_PS%" -AppPath "%APP_PATH%" -AppArgs "%APP_ARGS%" -AppName "FlowVision"

endlocal

