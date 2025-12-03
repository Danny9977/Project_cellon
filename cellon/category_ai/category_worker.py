# cellon/category_ai/category_worker.py
from pathlib import Path
from typing import Optional

import pandas as pd
from PyQt6.QtCore import QThread, pyqtSignal

from .category_loader import get_category_master


class CategoryBuildWorker(QThread):
    """
    카테고리 엑셀들을 읽어서 마스터 테이블을 만드는 워커(QThread).
    - 진행 상황은 progress(int, str) 시그널로 UI에 전달
    - 완료 시 finished(DataFrame) 시그널로 결과 전달
    - 오류 시 error(str) 시그널로 메시지 전달
    """
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(object)  # pd.DataFrame
    error = pyqtSignal(str)

    def __init__(self, category_dir: Path, parent=None):
        super().__init__(parent)
        self.category_dir = Path(category_dir)

    def run(self):
        try:
            def _cb(p: int, msg: str):
                # category_loader 에서 progress_cb 로 호출되면
                # 여기서 시그널로 UI에 넘김
                self.progress.emit(int(p), str(msg))

            df: pd.DataFrame = get_category_master(
                category_dir=self.category_dir,
                progress_cb=_cb
            )
            self.finished.emit(df)
        except Exception as e:
            self.error.emit(str(e))
