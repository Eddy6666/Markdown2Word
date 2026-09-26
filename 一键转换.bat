@echo off
rem ============================================
rem  md to Word one-click launcher
rem  Drag .md files onto this file to convert.
rem ============================================
where python >nul 2>nul
if errorlevel 1 (
    echo Python not found. Please install Python 3 and add it to PATH.
    pause
    exit /b 1
)
python "%~dp0md2word.py" %*
