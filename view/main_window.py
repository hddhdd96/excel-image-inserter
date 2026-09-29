"""
메인 GUI 창
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import Optional, Callable
import threading
import os
import sys

# 프로젝트 루트를 경로에 추가
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from model.logger import LogType


class MainWindow:
    """메인 윈도우 클래스"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("엑셀 이미지 삽입 프로그램")
        self.root.geometry("800x750")
        
        # 콜백 함수들
        self.on_execute: Optional[Callable] = None
        
        # 변수들
        self.excel_file_var = tk.StringVar()
        self.image_folder_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.match_column_var = tk.StringVar(value="A")
        self.start_cell_var = tk.StringVar(value="C5")  # B5 -> C5로 변경
        
        # 이미지 크기 설정 변수
        self.image_size_mode_var = tk.StringVar(value="fit_cell")
        self.image_width_var = tk.StringVar(value="")
        self.image_height_var = tk.StringVar(value="")
        
        self._create_widgets()
    
    def _create_widgets(self):
        """위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # 엑셀 파일 선택
        ttk.Label(main_frame, text="엑셀 파일:").grid(row=row, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_file_var, width=50).grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(main_frame, text="찾아보기", command=self._browse_excel_file).grid(row=row, column=2, pady=5)
        row += 1
        
        # 시트 선택
        ttk.Label(main_frame, text="시트:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.sheet_combo = ttk.Combobox(main_frame, textvariable=self.sheet_var, state="readonly", width=47)
        self.sheet_combo.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(main_frame, text="시트 새로고침", command=self._refresh_sheets).grid(row=row, column=2, pady=5)
        row += 1
        
        # 이미지 폴더 선택
        ttk.Label(main_frame, text="이미지 폴더:").grid(row=row, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.image_folder_var, width=50).grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(main_frame, text="찾아보기", command=self._browse_image_folder).grid(row=row, column=2, pady=5)
        row += 1
        
        # 매칭 기준 열
        ttk.Label(main_frame, text="매칭 기준 열:").grid(row=row, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.match_column_var, width=10).grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(main_frame, text="(예: A, B, C)").grid(row=row, column=1, sticky=tk.W, padx=(80, 0))
        row += 1
        
        # 이미지 삽입 시작 셀
        ttk.Label(main_frame, text="삽입 시작 셀:").grid(row=row, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.start_cell_var, width=10).grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(main_frame, text="(예: B5, C10)").grid(row=row, column=1, sticky=tk.W, padx=(80, 0))
        row += 1
        
        # 이미지 크기 설정
        ttk.Label(main_frame, text="이미지 크기:").grid(row=row, column=0, sticky=tk.W, pady=5)
        size_frame = ttk.Frame(main_frame)
        size_frame.grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Radiobutton(size_frame, text="셀 크기에 맞춤", variable=self.image_size_mode_var, 
                       value="fit_cell").grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Radiobutton(size_frame, text="수동 설정", variable=self.image_size_mode_var, 
                       value="manual").grid(row=0, column=1, sticky=tk.W)
        row += 1
        
        # 수동 크기 설정 입력 필드
        self.size_manual_frame = ttk.Frame(main_frame)
        self.size_manual_frame.grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(self.size_manual_frame, text="너비:").grid(row=0, column=0, padx=(20, 5))
        ttk.Entry(self.size_manual_frame, textvariable=self.image_width_var, width=8).grid(row=0, column=1, padx=5)
        ttk.Label(self.size_manual_frame, text="픽셀").grid(row=0, column=2, padx=5)
        
        ttk.Label(self.size_manual_frame, text="높이:").grid(row=0, column=3, padx=(20, 5))
        ttk.Entry(self.size_manual_frame, textvariable=self.image_height_var, width=8).grid(row=0, column=4, padx=5)
        ttk.Label(self.size_manual_frame, text="픽셀").grid(row=0, column=5, padx=5)
        
        # 라디오 버튼 변경 시 수동 입력 필드 표시/숨김
        self.image_size_mode_var.trace('w', self._on_size_mode_changed)
        self._on_size_mode_changed()  # 초기 상태 설정
        row += 1
        
        # 실행 버튼
        self.execute_button = ttk.Button(main_frame, text="실행", command=self._on_execute_clicked, state=tk.DISABLED)
        self.execute_button.grid(row=row, column=1, pady=20)
        row += 1
        
        # 로그 출력 영역
        ttk.Label(main_frame, text="로그:").grid(row=row, column=0, sticky=(tk.W, tk.N), pady=5)
        self.log_text = scrolledtext.ScrolledText(main_frame, width=70, height=20, wrap=tk.WORD)
        self.log_text.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        main_frame.rowconfigure(row, weight=1)
        row += 1
        
        # 초기 시트 새로고침
        self._refresh_sheets()
    
    def _on_size_mode_changed(self, *args):
        """이미지 크기 모드 변경 시 호출"""
        if self.image_size_mode_var.get() == "manual":
            self.size_manual_frame.grid()
        else:
            self.size_manual_frame.grid_remove()
    
    def _browse_excel_file(self):
        """엑셀 파일 선택 다이얼로그"""
        filename = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file_var.set(filename)
            self._refresh_sheets()
    
    def _browse_image_folder(self):
        """이미지 폴더 선택 다이얼로그"""
        folder = filedialog.askdirectory(title="이미지 폴더 선택")
        if folder:
            self.image_folder_var.set(folder)
    
    def _refresh_sheets(self):
        """시트 목록 새로고침"""
        excel_file = self.excel_file_var.get()
        if not excel_file:
            self.sheet_combo['values'] = []
            self.sheet_var.set("")
            return
        
        # 시트 목록을 가져오는 콜백이 있으면 호출
        if hasattr(self, 'on_get_sheets'):
            sheets = self.on_get_sheets(excel_file)
            if sheets:
                self.sheet_combo['values'] = sheets
                if sheets:
                    self.sheet_var.set(sheets[0])
    
    def _on_execute_clicked(self):
        """실행 버튼 클릭 이벤트"""
        if self.on_execute:
            # 별도 스레드에서 실행하여 UI가 멈추지 않도록
            thread = threading.Thread(target=self._execute_in_thread)
            thread.daemon = True
            thread.start()
    
    def _execute_in_thread(self):
        """별도 스레드에서 실행"""
        self.root.after(0, self._disable_execute_button)
        try:
            if self.on_execute:
                self.on_execute()
        finally:
            self.root.after(0, self._enable_execute_button)
    
    def _disable_execute_button(self):
        """실행 버튼 비활성화"""
        self.execute_button.config(state=tk.DISABLED)
    
    def _enable_execute_button(self):
        """실행 버튼 활성화"""
        self.execute_button.config(state=tk.NORMAL)
    
    def set_get_sheets_callback(self, callback: Callable[[str], list]):
        """시트 목록 가져오기 콜백 설정"""
        self.on_get_sheets = callback
    
    def set_execute_callback(self, callback: Callable):
        """실행 콜백 설정"""
        self.on_execute = callback
        self.execute_button.config(state=tk.NORMAL)
    
    def log(self, message: str, log_type: LogType = LogType.INFO):
        """
        로그 출력
        
        Args:
            message: 로그 메시지
            log_type: 로그 타입
        """
        # 색상 태그 설정
        tag = log_type.value.lower()
        if tag not in self.log_text.tag_names():
            colors = {
                'success': 'green',
                'fail': 'orange',
                'error': 'red',
                'summary': 'blue',
                'info': 'black',
                'warning': 'orange'
            }
            color = colors.get(tag, 'black')
            self.log_text.tag_config(tag, foreground=color)
        
        # 로그 추가
        self.log_text.insert(tk.END, message + "\n", tag)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log(self):
        """로그 초기화"""
        self.log_text.delete(1.0, tk.END)
    
    def show_error(self, title: str, message: str):
        """에러 메시지 박스 표시"""
        messagebox.showerror(title, message)
    
    def show_info(self, title: str, message: str):
        """정보 메시지 박스 표시"""
        messagebox.showinfo(title, message)
    
    def get_excel_file(self) -> str:
        """엑셀 파일 경로 반환"""
        return self.excel_file_var.get()
    
    def get_image_folder(self) -> str:
        """이미지 폴더 경로 반환"""
        return self.image_folder_var.get()
    
    def get_sheet_name(self) -> str:
        """선택된 시트 이름 반환"""
        return self.sheet_var.get()
    
    def get_match_column(self) -> str:
        """매칭 기준 열 반환"""
        return self.match_column_var.get().strip().upper()
    
    def get_start_cell(self) -> str:
        """시작 셀 반환"""
        return self.start_cell_var.get().strip().upper()
    
    def get_image_size_mode(self) -> str:
        """이미지 크기 모드 반환"""
        return self.image_size_mode_var.get()
    
    def get_image_width(self) -> Optional[int]:
        """이미지 너비 반환 (픽셀)"""
        width_str = self.image_width_var.get().strip()
        if width_str:
            try:
                return int(width_str)
            except ValueError:
                return None
        return None
    
    def get_image_height(self) -> Optional[int]:
        """이미지 높이 반환 (픽셀)"""
        height_str = self.image_height_var.get().strip()
        if height_str:
            try:
                return int(height_str)
            except ValueError:
                return None
        return None
    
    def run(self):
        """GUI 실행"""
        self.root.mainloop()
