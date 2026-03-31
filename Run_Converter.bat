@echo off
title W2PDFWord to PDF Bulk Converter
echo =========================================
echo Starting Word to PDF Bulk Converter...
echo =========================================
echo.
echo Please wait while the local server starts...

:: Change directory to where the batch file is located
cd /d "%~dp0"

:: Run the Streamlit application
streamlit run app.py

:: Pause in case of an error so the window doesn't close immediately
pause
