"""
에러 코드 정의
"""

from enum import Enum


class ErrorCode(Enum):
    """에러 코드 타입"""
    E001 = "엑셀 로드 실패"
    E002 = "이미지 폴더 없음"
    E003 = "셀 좌표 오류"
    E004 = "이미지 삽입 실패"
    
    def __str__(self):
        return f"{self.name}: {self.value}"
