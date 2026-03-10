@echo off
REM Markdown to Word Converter - Application Launcher
REM This script starts the Flask application

cd /d "%~dp0"
echo.
echo ============================================================
echo   MARKDOWN TO WORD CONVERTER
echo ============================================================
echo   Starting Flask server...
echo   Open browser: http://localhost:5000
echo ============================================================
echo.

python server.py

pause
