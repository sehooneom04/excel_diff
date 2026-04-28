# filepath: excel_diff_app/ui/main_window.py
"""메인 윈도우 UI"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import sys

from ..core.differ import diff_excel, get_total_changes, format_stats_message
from ..core.constants import DEFAULT_OUTPUT, FILE_FILTERS


class MainWindow:
    """메인 윈도우 클래스"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 파일 비교")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # 파일 경로 저장
        self.file1_path = None
        self.file2_path = None
        self.output_path = None
        
        self._setup_ui()
    
    def _setup_ui(self):
        """UI 구성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목
        title_label = ttk.Label(
            main_frame, 
            text="Excel 파일 비교 도구",
            font=("Arial", 18, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 파일 선택 프레임
        file_frame = ttk.LabelFrame(main_frame, text="파일 선택", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        
        # 원본 파일
        ttk.Label(file_frame, text="원본 파일:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.file1_entry = ttk.Entry(file_frame, width=50)
        self.file1_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="찾아보기...", command=self._select_file1).grid(row=0, column=2, pady=5)
        
        # 수정본 파일
        ttk.Label(file_frame, text="수정본 파일:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.file2_entry = ttk.Entry(file_frame, width=50)
        self.file2_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="찾아보기...", command=self._select_file2).grid(row=1, column=2, pady=5)
        
        # 저장 위치
        ttk.Label(file_frame, text="저장 위치:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_entry = ttk.Entry(file_frame, width=50)
        self.output_entry.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="찾아보기...", command=self._select_output).grid(row=2, column=2, pady=5)
        
        # 기본 출력 파일명 설정
        self.output_entry.insert(0, DEFAULT_OUTPUT)
        
        # 진행률 표시
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_label = ttk.Label(self.progress_frame, text="")
        self.progress_label.pack()
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='indeterminate')
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        # 실행 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        self.run_button = ttk.Button(
            button_frame, 
            text="비교 실행", 
            command=self._run_comparison,
            width=15
        )
        self.run_button.pack()
        
        # 상태 메시지
        self.status_label = ttk.Label(main_frame, text="파일을 선택해주세요", foreground="gray")
        self.status_label.pack(pady=10)
    
    def _select_file1(self):
        """원본 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="원본 Excel 파일 선택",
            filetypes=FILE_FILTERS
        )
        if file_path:
            self.file1_path = file_path
            self.file1_entry.delete(0, tk.END)
            self.file1_entry.insert(0, file_path)
            self._update_status()
    
    def _select_file2(self):
        """수정본 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="수정본 Excel 파일 선택",
            filetypes=FILE_FILTERS
        )
        if file_path:
            self.file2_path = file_path
            self.file2_entry.delete(0, tk.END)
            self.file2_entry.insert(0, file_path)
            self._update_status()
    
    def _select_output(self):
        """저장 위치 선택"""
        file_path = filedialog.asksaveasfilename(
            title="결과 파일 저장 위치",
            defaultextension=".xlsx",
            filetypes=FILE_FILTERS
        )
        if file_path:
            self.output_path = file_path
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)
    
    def _update_status(self):
        """상태 메시지 업데이트"""
        if self.file1_path and self.file2_path:
            f1 = Path(self.file1_path).name
            f2 = Path(self.file2_path).name
            self.status_label.config(text=f"{f1} vs {f2}")
        else:
            self.status_label.config(text="파일을 선택해주세요", foreground="gray")
    
    def _run_comparison(self):
        """비교 실행"""
        # 파일 검증
        if not self.file1_path:
            messagebox.showerror("오류", "원본 파일을 선택해주세요")
            return
        
        if not self.file2_path:
            messagebox.showerror("오류", "수정본 파일을 선택해주세요")
            return
        
        if not Path(self.file1_path).exists():
            messagebox.showerror("오류", f"원본 파일이 없습니다: {self.file1_path}")
            return
        
        if not Path(self.file2_path).exists():
            messagebox.showerror("오류", f"수정본 파일이 없습니다: {self.file2_path}")
            return
        
        # 출력 경로
        output = self.output_entry.get().strip()
        if not output:
            output = DEFAULT_OUTPUT
        
        # UI 비활성화
        self.run_button.config(state=tk.DISABLED)
        self.progress_bar.start(10)
        
        try:
            # 비교 실행
            stats = diff_excel(
                Path(self.file1_path),
                Path(self.file2_path),
                Path(output)
            )
            
            # 결과 표시
            total = get_total_changes(stats)
            message = format_stats_message(stats)
            messagebox.showinfo(
                "완료",
                f"✅ 비교 완료!\n\n총 {total}개 변경\n\n{message}\n\n결과 파일: {output}"
            )
            
        except Exception as e:
            messagebox.showerror("오류", f"비교 중 오류가 발생했습니다:\n{str(e)}")
        
        finally:
            self.progress_bar.stop()
            self.run_button.config(state=tk.NORMAL)


def run_gui():
    """GUI 실행"""
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()


if __name__ == "__main__":
    run_gui()