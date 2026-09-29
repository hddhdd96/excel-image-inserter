"""
Presenter 클래스
"""

import os
import sys
from typing import Optional

# 프로젝트 루트를 경로에 추가
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from model import ExcelHandler, ImageMatcher, ImageProcessor, Logger, ErrorCode
from view import MainWindow


class Presenter:
    """View와 Model을 연결하는 Presenter"""
    
    def __init__(self):
        self.view = MainWindow()
        self.logger = Logger(callback=self._on_log)
        
        # Model 인스턴스들
        self.excel_handler: Optional[ExcelHandler] = None
        self.image_matcher: Optional[ImageMatcher] = None
        self.image_processor: Optional[ImageProcessor] = None
        
        # 콜백 설정
        self.view.set_get_sheets_callback(self._get_sheets)
        self.view.set_execute_callback(self._execute)
    
    def _on_log(self, message: str, log_type):
        """로그 콜백"""
        self.view.log(message, log_type)
    
    def _get_sheets(self, excel_file: str) -> list:
        """
        시트 목록 가져오기
        
        Args:
            excel_file: 엑셀 파일 경로
            
        Returns:
            시트 이름 리스트
        """
        try:
            if not excel_file or not os.path.exists(excel_file):
                return []
            
            handler = ExcelHandler(excel_file, self.logger)
            if handler.load():
                sheets = handler.get_sheet_names()
                handler.close()
                return sheets
            return []
        except Exception as e:
            self.logger.error(f"시트 목록 가져오기 실패: {str(e)}")
            return []
    
    def _validate_inputs(self) -> bool:
        """입력값 검증"""
        excel_file = self.view.get_excel_file()
        image_folder = self.view.get_image_folder()
        sheet_name = self.view.get_sheet_name()
        match_column = self.view.get_match_column()
        start_cell = self.view.get_start_cell()
        
        # 엑셀 파일 검증
        if not excel_file:
            self.view.show_error("입력 오류", "엑셀 파일을 선택해주세요.")
            return False
        
        if not os.path.exists(excel_file):
            self.view.show_error("입력 오류", f"엑셀 파일이 존재하지 않습니다: {excel_file}")
            return False
        
        # 이미지 폴더 검증
        if not image_folder:
            self.view.show_error("입력 오류", "이미지 폴더를 선택해주세요.")
            return False
        
        if not os.path.exists(image_folder):
            self.view.show_error("입력 오류", f"이미지 폴더가 존재하지 않습니다: {image_folder}")
            return False
        
        if not os.path.isdir(image_folder):
            self.view.show_error("입력 오류", f"이미지 폴더가 디렉토리가 아닙니다: {image_folder}")
            return False
        
        # 시트 검증
        if not sheet_name:
            self.view.show_error("입력 오류", "시트를 선택해주세요.")
            return False
        
        # 매칭 기준 열 검증
        if not match_column:
            self.view.show_error("입력 오류", "매칭 기준 열을 입력해주세요.")
            return False
        
        if len(match_column) > 2 or not match_column[0].isalpha():
            self.view.show_error("입력 오류", "매칭 기준 열은 A-Z 또는 AA-ZZ 형식이어야 합니다.")
            return False
        
        # 시작 셀 검증
        if not start_cell:
            self.view.show_error("입력 오류", "삽입 시작 셀을 입력해주세요.")
            return False
        
        # 셀 형식 검증 (예: B5, AA10)
        if len(start_cell) < 2:
            self.view.show_error("입력 오류", "시작 셀 형식이 올바르지 않습니다. (예: B5)")
            return False
        
        # 열 부분과 행 부분 분리
        col_part = ""
        row_part = ""
        for i, char in enumerate(start_cell):
            if char.isalpha():
                col_part += char
            else:
                row_part = start_cell[i:]
                break
        
        if not col_part or not row_part or not row_part.isdigit():
            self.view.show_error("입력 오류", "시작 셀 형식이 올바르지 않습니다. (예: B5)")
            return False
        
        return True
    
    def _execute(self):
        """실행 처리"""
        # 로그 초기화
        self.view.clear_log()
        self.logger.clear()
        
        # 입력 검증
        if not self._validate_inputs():
            return
        
        try:
            excel_file = self.view.get_excel_file()
            image_folder = self.view.get_image_folder()
            sheet_name = self.view.get_sheet_name()
            match_column = self.view.get_match_column()
            start_cell = self.view.get_start_cell()
            
            self.logger.info("=" * 50)
            self.logger.info("이미지 삽입 작업 시작")
            self.logger.info("=" * 50)
            
            # ExcelHandler 생성 및 로드
            self.excel_handler = ExcelHandler(excel_file, self.logger)
            if not self.excel_handler.load():
                self.view.show_error("오류", f"{ErrorCode.E001}: 엑셀 파일을 로드할 수 없습니다.")
                return
            
            # 시트 선택
            if not self.excel_handler.select_sheet(sheet_name):
                self.view.show_error("오류", f"시트를 선택할 수 없습니다: {sheet_name}")
                self.excel_handler.close()
                return
            
            # ImageMatcher 생성
            self.image_matcher = ImageMatcher(image_folder, self.logger)
            
            # ImageProcessor 생성
            self.image_processor = ImageProcessor(self.excel_handler, self.image_matcher, self.logger)
            
            # 이미지 크기 설정 가져오기
            image_size_mode = self.view.get_image_size_mode()
            fit_to_cell = (image_size_mode == "fit_cell")
            image_width = self.view.get_image_width() if not fit_to_cell else None
            image_height = self.view.get_image_height() if not fit_to_cell else None
            
            # 수동 설정 모드일 때 크기 값 검증
            if not fit_to_cell and (image_width is None and image_height is None):
                self.logger.warning("수동 설정 모드이지만 크기가 입력되지 않았습니다. 원본 크기로 삽입됩니다.")
            
            # 이미지 삽입 처리
            total_rows, row_success, row_fail, total_images, image_success, image_fail = self.image_processor.process(
                match_column, start_cell, 
                fit_to_cell=fit_to_cell,
                width=image_width, 
                height=image_height
            )
            
            # 엑셀 파일 저장
            if image_success > 0:
                if self.excel_handler.save():
                    self.logger.info("=" * 50)
                    self.logger.summary(f"작업 완료!")
                    self.view.show_info("완료", 
                                       f"작업이 완료되었습니다.\n\n"
                                       f"[행 단위]\n"
                                       f"총 {total_rows}행 중 {row_success}개 성공, {row_fail}개 실패\n\n"
                                       f"[이미지 단위]\n"
                                       f"총 {total_images}개 중 {image_success}개 성공, {image_fail}개 실패")
                else:
                    self.view.show_error("오류", "엑셀 파일 저장에 실패했습니다.")
            else:
                self.logger.warning("삽입된 이미지가 없습니다.")
                self.view.show_info("알림", "삽입된 이미지가 없습니다.")
            
            # 정리
            self.excel_handler.close()
            
        except Exception as e:
            self.logger.error(f"예기치 않은 오류 발생: {str(e)}")
            self.view.show_error("오류", f"예기치 않은 오류가 발생했습니다:\n{str(e)}")
        finally:
            # 정리
            if self.excel_handler:
                self.excel_handler.close()
                self.excel_handler = None
    
    def run(self):
        """프로그램 실행"""
        self.view.run()
