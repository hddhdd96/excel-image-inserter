"""
로거 클래스
"""

from enum import Enum
from typing import Callable, Optional


class LogType(Enum):
    """로그 타입"""
    SUCCESS = "SUCCESS"
    FAIL = "FAIL"
    ERROR = "ERROR"
    SUMMARY = "SUMMARY"
    INFO = "INFO"
    WARNING = "WARNING"


class Logger:
    """로그 출력을 위한 클래스"""
    
    def __init__(self, callback: Optional[Callable[[str, LogType], None]] = None):
        """
        Args:
            callback: 로그를 출력할 콜백 함수 (message, log_type) -> None
        """
        self.callback = callback
        self.logs = []
    
    def log(self, message: str, log_type: LogType = LogType.INFO):
        """
        로그 메시지 기록
        
        Args:
            message: 로그 메시지
            log_type: 로그 타입
        """
        log_entry = f"[{log_type.value}] {message}"
        self.logs.append((log_entry, log_type))
        
        if self.callback:
            self.callback(log_entry, log_type)
    
    def success(self, message: str):
        """성공 로그"""
        self.log(message, LogType.SUCCESS)
    
    def fail(self, message: str):
        """실패 로그"""
        self.log(message, LogType.FAIL)
    
    def error(self, message: str):
        """에러 로그"""
        self.log(message, LogType.ERROR)
    
    def summary(self, message: str):
        """요약 로그"""
        self.log(message, LogType.SUMMARY)
    
    def info(self, message: str):
        """정보 로그"""
        self.log(message, LogType.INFO)
    
    def warning(self, message: str):
        """경고 로그"""
        self.log(message, LogType.WARNING)
    
    def clear(self):
        """로그 초기화"""
        self.logs.clear()
