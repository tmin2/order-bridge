#!/bin/bash
# OrderBridge Dashboard Shutdown for macOS
cd "$(dirname "$0")"

clear
echo "========================================================"
echo "  OrderBridge 대시보드 종료기 (macOS)"
echo "========================================================"
echo ""
echo "실행 중인 대시보드 서버를 안전하게 종료합니다..."

PID_8501=$(lsof -ti:8501 2>/dev/null)
PID_8502=$(lsof -ti:8502 2>/dev/null)

if [ -n "$PID_8501" ]; then
    kill -9 $PID_8501 2>/dev/null
fi

if [ -n "$PID_8502" ]; then
    kill -9 $PID_8502 2>/dev/null
fi

echo ""
echo "✅ 대시보드가 안전하게 종료되었습니다."
sleep 2