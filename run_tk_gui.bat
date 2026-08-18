@echo off
setlocal

cd /d "%~dp0"
if exist "book-reader.exe" (
    start "" "book-reader.exe"
) else if exist "dist\book-reader.exe" (
    start "" "dist\book-reader.exe"
) else (
    python book_to_latex_gui.py
)
