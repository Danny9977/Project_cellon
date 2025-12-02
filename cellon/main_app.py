# main_app.py
import sys
from PyQt6.QtWidgets import QApplication

from ui_main import ChromeCrawler  # 👉 UI/로직은 모두 ui_main.py 안에 있음

# === 카테고리 AI 마스터 데이터 로드 ===
#from cellon.category_ai.category_loader import build_category_master

# 앱 초기화 시 한 번만 카테고리 마스터 로드
#CATEGORY_MASTER_DF = build_category_master()
# ==================================



if __name__ == "__main__":
    print(">>> main_app start")
    app = QApplication(sys.argv)
    win = ChromeCrawler()
    win.show()
    print(">>> before app.exec()")
    sys.exit(app.exec())
