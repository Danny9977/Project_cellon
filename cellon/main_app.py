# main_app.py
import sys
from PyQt6.QtWidgets import QApplication
from cellon.ui_main import ChromeCrawler   # ← 패키지 경로 명시
import os

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
sys.path.append(PROJECT_ROOT)

def main() -> None:
    app = QApplication(sys.argv)
    win = ChromeCrawler()
    win.show()
    app.exec()

# === 카테고리 AI 마스터 데이터 로드 ===
#from cellon.category_ai.category_loader import build_category_master

# 앱 초기화 시 한 번만 카테고리 마스터 로드
#CATEGORY_MASTER_DF = build_category_master()
# ==================================



if __name__ == "__main__":
    print(">>> main_app start")
    main()

