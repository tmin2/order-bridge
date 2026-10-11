#!/bin/bash
# OrderBridge Dashboard Launcher for macOS (MacBook)
cd "$(dirname "$0")"

clear
echo "========================================================"
echo "  OrderBridge 발주서 변환 대시보드 (macOS)"
echo "========================================================"
echo ""

# 1. 파이썬 확인
if command -v python3 &>/dev/null; then
    PY_BIN="python3"
elif command -v python &>/dev/null; then
    PY_BIN="python"
else
    echo "❌ Python(파이썬 3)이 설치되어 있지 않습니다."
    echo "   맥북에 파이썬을 설치하신 후 다시 실행해주세요."
    echo "   다운로드: https://www.python.org/downloads/"
    echo ""
    read -p "엔터(Enter) 키를 누르면 종료됩니다..."
    exit 1
fi

# 2. 맥북 전용 가상환경(.venv_mac) 구성 (macOS 외부 패키지 충돌 방지)
VENV_DIR=".venv_mac"
if [ ! -d "$VENV_DIR" ]; then
    echo "⚙️  최초 1회 맥북 전용 환경을 구성하고 있습니다. 잠시만 기다려주세요..."
    $PY_BIN -m venv "$VENV_DIR"
fi

# 가상환경 활성화
if [ -f "$VENV_DIR/bin/activate" ]; then
    source "$VENV_DIR/bin/activate"
    RUN_PY="python"
else
    RUN_PY="$PY_BIN"
fi

# 3. 필수 패키지 설치 확인
if ! $RUN_PY -c "import streamlit, pandas, openpyxl, xlrd" &>/dev/null; then
    echo "📦 필수 패키지(Streamlit, Pandas 등)를 설치하고 있습니다 (최초 1회)..."
    $RUN_PY -m pip install --upgrade pip
    $RUN_PY -m pip install -r requirements.txt
fi

echo ""
echo "🚀 대시보드 서버 및 브라우저를 실행합니다..."
echo "   종료하려면 터미널 창을 닫거나 Ctrl+C를 누르세요."
echo ""

$RUN_PY -m streamlit run app.py