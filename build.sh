#!/bin/bash
# filepath: build.sh
# Excel Diff Application Build Script for macOS/Linux

echo ""
echo "========================================"
echo "Excel Diff Application 빌드 (macOS)"
echo "========================================"
echo ""

# Python 환경 확인
if ! command -v python3 &> /dev/null; then
    echo "오류: Python가 설치되어 있지 않습니다."
    echo "Python 3.8 이상을 설치해주세요."
    exit 1
fi

# 필요한 패키지 설치
echo "[1/4] 필요한 패키지 설치 중..."
pip3 install openpyxl pyinstaller --quiet

# 빌드 디렉토리 생성
mkdir -p dist

# PyInstaller로 빌드
echo "[2/4] 빌드 중..."
pyinstaller excel_diff.spec --clean

echo ""
echo "[3/4] 빌드 결과 확인..."
if [ -f "dist/ExcelDiff" ]; then
    echo ""
    echo "========================================"
    echo "✅ 빌드 완료!"
    echo "========================================"
    echo ""
    echo "생성된 파일: dist/ExcelDiff"
    echo ""
    echo "실행 방법: ./dist/ExcelDiff"
    echo ""
else
    echo ""
    echo "오류: 빌드에 실패했습니다."
    echo ""
    exit 1
fi

echo ""
echo "모든 작업이 완료되었습니다."