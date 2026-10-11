@echo off
cd /d "%~dp0"
cls
title OrderBridge 대시보드 실행

echo ========================================================
echo   OrderBridge 발주서 변환 대시보드 실행
echo ========================================================
echo.
echo   서버 및 브라우저를 구동 중입니다. 잠시만 기다려주세요...
echo.

if exist "%~dp0python_runtime\python.exe" (
    "%~dp0python_runtime\python.exe" -m streamlit run app.py
) else (
    python -m streamlit run app.py
)

pause