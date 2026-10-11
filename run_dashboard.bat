@echo off
cd /d "%~dp0"
cls
title OrderBridge Dashboard

echo ========================================================
echo   OrderBridge Dashboard
echo ========================================================
echo.
echo   Starting server and browser, please wait...
echo.

if exist "%~dp0python_runtime\python.exe" (
    "%~dp0python_runtime\python.exe" -m streamlit run app.py
) else (
    python -m streamlit run app.py
)

pause