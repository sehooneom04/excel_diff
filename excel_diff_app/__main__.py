# filepath: excel_diff_app/__main__.py
"""
Excel 파일 비교 애플리케이션

사용법:
    GUI 모드: python -m excel_diff_app
    명령행: python -m excel_diff_app 원본.xlsx 수정본.xlsx [-o 결과.xlsx]
"""

import sys
import argparse
from pathlib import Path

# GUI 모듈 import
from .ui.main_window import run_gui
from .core.differ import diff_excel, get_total_changes, format_stats_message
from .core.constants import DEFAULT_OUTPUT


def main_cli():
    """명령행 인터페이스"""
    parser = argparse.ArgumentParser(
        description="두 엑셀 파일 비교 — 수정본 서식 그대로 유지하며 변경 셀만 표시"
    )
    parser.add_argument("file1", help="원본 파일 (.xlsx)")
    parser.add_argument("file2", help="수정본 파일 (.xlsx)")
    parser.add_argument("-o", "--output", default=None,
                        help="출력 파일명 (기본값: diff_result.xlsx)")
    args = parser.parse_args()

    f1 = Path(args.file1)
    f2 = Path(args.file2)
    out = Path(args.output) if args.output else Path(DEFAULT_OUTPUT)

    # 파일 존재 확인
    if not f1.exists():
        print(f"오류: {f1} 파일이 없습니다")
        sys.exit(1)
    if not f2.exists():
        print(f"오류: {f2} 파일이 없습니다")
        sys.exit(1)

    print(f"비교: {f1.name} vs {f2.name}")
    
    try:
        stats = diff_excel(f1, f2, out)
        total = get_total_changes(stats)
        
        print(f"\n{format_stats_message(stats)}")
        print(f"\n✅ 완료: {out}  (총 {total}개 변경)")
        
    except Exception as e:
        print(f"오류: {str(e)}")
        sys.exit(1)


def main():
    """메인 진입점"""
    if len(sys.argv) == 1:
        # 인자 없이 실행 시 GUI 모드
        run_gui()
    else:
        # 인자가 있으면 명령행 모드
        main_cli()


if __name__ == "__main__":
    main()