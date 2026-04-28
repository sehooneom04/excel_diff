@echo off
REM ============================================
REM Excel Diff Application Build Script
REM ============================================

echo.
echo ========================================
echo Excel Diff Application 빌드
echo ========================================
echo.

REM Python 환경 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo 오류: Python가 설치되어 있지 않습니다.
    echo Python 3.8 이상을 설치해주세요.
    pause
    exit /b 1
)

REM 필요한 패키지 설치
echo [1/4] 필요한 패키지 설치 중...
pip install openpyxl pyinstaller --quiet

REM 빌드 디렉토리 생성
if not exist "dist\ExcelDiff" mkdir "dist\ExcelDiff"

REM PyInstaller로 빌드
echo [2/4] 빌드 중...
pyinstaller excel_diff.spec --clean

echo.
echo [3/4] 빌드 결과 확인...
if exist "dist\ExcelDiff.exe" (
    echo.
    echo ========================================
    echo ✅ 빌드 완료!
    echo ========================================
    echo.
    echo 생성된 파일: dist\ExcelDiff.exe
    echo.
    echo 실행 방법: dist\ExcelDiff.exe
    echo.
) else (
    echo.
    echo 오류: 빌드에 실패했습니다.
    echo.
    pause
    exit /b 1
)

REM [4/4] 정리
echo [4/4] 임시 파일 정리 중...
if exist "*.spec~" del "*.spec~" 2>nul

echo.
echo 모든 작업이 완료되었습니다.
pause