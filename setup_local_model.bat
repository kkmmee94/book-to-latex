@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_local_model.ps1"
if errorlevel 1 (
    echo.
    echo Local model setup failed. Read the messages above.
    pause
    exit /b 1
)
echo.
echo Local model setup finished.
pause
