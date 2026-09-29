"""
메인 진입점
"""

import sys
import os

# 프로젝트 루트를 Python 경로에 추가
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# 패키지 이름을 사용한 절대 import
from presenter.presenter import Presenter


def main():
    """메인 함수"""
    presenter = Presenter()
    presenter.run()


if __name__ == "__main__":
    main()
