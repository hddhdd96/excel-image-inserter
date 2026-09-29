"""
이미지 파일 매칭 클래스
"""

import os
from typing import List
from .logger import Logger


class ImageMatcher:
    """이미지 파일 매칭 처리"""
    
    def __init__(self, image_folder: str, logger: Logger = None):
        """
        Args:
            image_folder: 이미지 폴더 경로
            logger: 로거 인스턴스
        """
        self.image_folder = image_folder
        self.logger = logger or Logger()
        self.image_files = []
        self._load_images()
    
    def _load_images(self):
        """이미지 폴더에서 파일 목록 로드"""
        if not os.path.exists(self.image_folder):
            self.logger.error(f"이미지 폴더가 존재하지 않습니다: {self.image_folder}")
            return
        
        if not os.path.isdir(self.image_folder):
            self.logger.error(f"이미지 폴더가 디렉토리가 아닙니다: {self.image_folder}")
            return
        
        # 지원하는 이미지 확장자
        image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif'}
        
        try:
            for filename in os.listdir(self.image_folder):
                file_path = os.path.join(self.image_folder, filename)
                if os.path.isfile(file_path):
                    _, ext = os.path.splitext(filename.lower())
                    if ext in image_extensions:
                        self.image_files.append(filename)
            
            self.logger.info(f"이미지 파일 {len(self.image_files)}개 로드 완료")
        except Exception as e:
            self.logger.error(f"이미지 폴더 읽기 실패: {str(e)}")
    
    def find_matching_images(self, cell_value: str) -> List[str]:
        """
        셀 값과 정확히 일치하거나 구분자 뒤에 오는 이미지 파일 찾기
        
        Args:
            cell_value: 셀 값 (예: '1', 'A12')
            
        Returns:
            매칭된 이미지 파일명 리스트 (전체 경로)
        """
        if not cell_value:
            return []
        
        matching_files = []
        cell_value_str = str(cell_value).strip()
        
        for filename in self.image_files:
            # 확장자를 제외한 파일명
            name_without_ext = os.path.splitext(filename)[0]
            
            # 정확히 일치하는 경우 (예: "1.png")
            if name_without_ext == cell_value_str:
                file_path = os.path.join(self.image_folder, filename)
                matching_files.append(file_path)
            # prefix로 시작하는 경우
            elif name_without_ext.startswith(cell_value_str) and \
                 len(name_without_ext) > len(cell_value_str):
                next_char = name_without_ext[len(cell_value_str)]
                
                # 셀 값이 숫자로만 이루어진 경우: 숫자가 아닌 모든 문자를 구분자로 인정
                # (예: "8대시보드-1.png", "8asdf.png", "1-a.png")
                if cell_value_str.isdigit() and not next_char.isdigit():
                    file_path = os.path.join(self.image_folder, filename)
                    matching_files.append(file_path)
                # 그 외의 경우: 특정 구분자만 인정
                # (예: "A12-front.png", "ABC_test.png")
                elif next_char in ['-', '_', ' ', '(', '.']:
                    file_path = os.path.join(self.image_folder, filename)
                matching_files.append(file_path)
        
        # 파일명 순서대로 정렬
        matching_files.sort()
        
        return matching_files
