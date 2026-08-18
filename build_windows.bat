@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_windows.ps1"
if errorlevel 1 (
    echo.
    echo Build failed. Read the messages above.
    pause
    exit /b 1
)
echo.
echo Build finished successfully.
pause
