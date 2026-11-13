@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" "%~dp0launch.py"
) else (
    python "%~dp0launch.py"
)
if errorlevel 1 (
    echo Please check the startup message above.
    pause
)
