# cellon/category_ai/category_worker.py
from PyQt6.QtCore import QThread, pyqtSignal  # PyQt6에서 스레드와 시그널 기능 가져옴
from pathlib import Path  # 파일/폴더 경로를 쉽게 다루는 모듈
from .category_loader import get_category_master  # 카테고리 마스터 생성 함수 import

class CategoryBuildWorker(QThread):  # PyQt6의 QThread를 상속받아 작업자 클래스 정의
    progress = pyqtSignal(int, str)   # 진행률(퍼센트, 메시지) 알림용 시그널
    finished = pyqtSignal(object)     # 작업 완료 시 결과(dataframe) 알림용 시그널

    def __init__(self, category_dir: Path):  # 생성자, 카테고리 폴더 경로 받음
        super().__init__()  # 부모(QThread) 초기화
        self.category_dir = category_dir  # 폴더 경로를 멤버 변수에 저장

    def run(self):  # 스레드에서 실행될 실제 작업 함수
        def _cb(p, m):  # 진행상황 콜백 함수, p: 퍼센트, m: 메시지
            self.progress.emit(p, m)  # 진행률 시그널로 알림

        try:  # 예외 처리 시작
            df = get_category_master(  # 카테고리 마스터 데이터 생성
                category_dir=self.category_dir,  # 폴더 경로 전달
                progress_cb=_cb  # 진행상황 콜백 전달
            )
            self.finished.emit(df)  # 작업 완료 시 결과(dataframe) 시그널로 알림
        except Exception as e:  # 오류 발생 시
            self.progress.emit(100, f"❌ 오류 발생: {e}")  # 오류 메시지 진행률로 알림
            self.finished.emit(None)  #
