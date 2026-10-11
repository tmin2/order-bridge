@echo off
cls
title OrderBridge 대시보드 종료

echo ========================================================
echo   OrderBridge 대시보드 종료기
echo ========================================================
echo.
echo 백그라운드에서 실행 중인 대시보드 서버를 안전하게 종료합니다...

:: 8501, 8502 포트를 점유하고 있는 프로세스 종료
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8501 ^| findstr LISTENING') do taskkill /f /pid %%a >nul 2>&1
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8502 ^| findstr LISTENING') do taskkill /f /pid %%a >nul 2>&1

echo.
echo [완료] 대시보드가 안전하게 종료되었습니다.
echo 2초 후 창이 자동으로 닫힙니다.
timeout /t 2 >nul