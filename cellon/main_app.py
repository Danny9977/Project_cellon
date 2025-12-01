# main_app.py
import sys
from PyQt6.QtWidgets import QApplication

from ui_main import ChromeCrawler  # 👉 UI/로직은 모두 ui_main.py 안에 있음


if __name__ == "__main__":
    print(">>> main_app start")
    app = QApplication(sys.argv)
    win = ChromeCrawler()
    win.show()
    print(">>> before app.exec()")
    sys.exit(app.exec())
