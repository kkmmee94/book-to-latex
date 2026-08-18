@echo off
setlocal
cd /d "%~dp0"
if "%~1"=="" (
    echo Drag the Qwen3-8B-hf folder onto this file, or run:
    echo setup_exact_local_model.bat "C:\path\to\Qwen3-8B-hf"
    pause
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_exact_local_model.ps1" -Weights "%~1"
if errorlevel 1 (
    echo.
    echo Exact local model setup failed. Read the messages above.
    pause
    exit /b 1
)
echo.
echo Exact local model setup finished.
pause
