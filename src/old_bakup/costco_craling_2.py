# 누끼는 따지 않는게 더 좋을 듯. 아니면 누끼를 약하게 조절해서 따게 하는게 좋을 듯 하얀색 다 날라감

import sys, os, io, base64
from datetime import datetime
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLineEdit, QPushButton, QTextEdit
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
import cv2
import numpy as np

class CrawlerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("코스트코 크롤러")
        self.layout = QVBoxLayout(self)

        self.url_input = QLineEdit(self)
        self.url_input.setPlaceholderText("코스트코 상품 URL 입력")
        self.layout.addWidget(self.url_input)

        self.log = QTextEdit(self)
        self.log.setReadOnly(True)
        self.layout.addWidget(self.log)

        self.btn_start = QPushButton("크롤링시작", self)
        self.btn_start.clicked.connect(self.start_crawl)
        self.layout.addWidget(self.btn_start)

        self.btn_send = QPushButton("data 보내기", self)
        self.btn_send.clicked.connect(self.save_to_numbers)
        self.layout.addWidget(self.btn_send)

        self.btn_stop = QPushButton("크롤링 스탑", self)
        self.btn_stop.clicked.connect(self.stop_crawl)
        self.layout.addWidget(self.btn_stop)

        self.crawled = []
        self.running = False

        # Selenium 드라이버 세팅
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        self.driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    def log_msg(self, msg):
        self.log.append(msg)

    def start_crawl(self):
        self.running = True
        url = self.url_input.text().strip()
        if not url:
            self.log_msg("⚠️ URL을 입력해주세요.")
            return
        self.log_msg(f"🔄 크롤링 시작: {url}")
        try:
            self.driver.get(url)
            wait = WebDriverWait(self.driver, 15)
            # 페이지 로드 대기
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.product-name")))

            # 이름과 가격 추출
            name_el = self.driver.find_element(By.CSS_SELECTOR, "h1.product-name")
            price_el = self.driver.find_element(By.CSS_SELECTOR, "span.notranslate.ng-star-inserted")
            name = name_el.text.strip()
            price = price_el.text.strip()
            self.crawled.append((name, price))
            self.log_msg(f"✅ 이름: {name}")
            self.log_msg(f"✅ 가격: {price}")

            # 이미지 캡쳐 및 누끼 제거
            self.capture_image(name)

        except Exception as e:
            self.log_msg(f"❌ 오류 발생: {e}")

    def capture_image(self, name):
        try:
            wait = WebDriverWait(self.driver, 10)
            img_el = wait.until(EC.presence_of_element_located((By.XPATH,
                '/html/body/main/div[4]/sip-product-details-page/sip-product-details/div/'
                'sip-product-image-panel/div/div/div[1]/div[1]/div[1]/div/'
                'sip-image-zoom/div/sip-media[1]/picture/img')))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", img_el)
            wait.until(EC.visibility_of(img_el))

            # 스크린샷 캡쳐
            screenshot = self.driver.get_screenshot_as_base64()
            image = Image.open(io.BytesIO(base64.b64decode(screenshot)))

            loc = img_el.location_once_scrolled_into_view
            size = img_el.size
            left, top = int(loc['x']), int(loc['y'])
            right, bottom = left + int(size['width']), top + int(size['height'])
            cropped = image.crop((left, top, right, bottom))

            # 누끼 제거: 배경 흰색으로 가정, OpenCV 활용
            cv_img = cv2.cvtColor(np.array(cropped), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
            _, thresh = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)
            contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            mask = np.zeros_like(cv_img)
            cv2.drawContours(mask, contours, -1, (255,255,255), cv2.FILLED)
            result = cv2.bitwise_and(cv_img, mask)
            # 배경 투명 PNG후 jpg로 저장(hard white bg)
            result = cv2.cvtColor(result, cv2.COLOR_BGR2RGB)
            pil = Image.fromarray(result)
            save_dir = self.get_save_dir()
            os.makedirs(save_dir, exist_ok=True)
            fname = os.path.join(save_dir, f"{name}.jpg")
            pil.save(fname, "JPEG")
            self.log_msg(f"✅ 이미지 저장 성공: {fname}")

        except Exception as e:
            self.log_msg(f"❌ 이미지 캡쳐 오류: {e}")

    def get_save_dir(self):
        today = datetime.now().strftime("%Y%m%d")
        path = os.path.expanduser(f"~/desktop/Python_Project/코스트코상품_{today}")
        return path

    def save_to_numbers(self):
        if not self.crawled:
            self.log_msg("⚠️ 저장할 데이터가 없습니다.")
            return
        try:
            import csv
            save_dir = self.get_save_dir()
            os.makedirs(save_dir, exist_ok=True)
            fname = os.path.join(save_dir, os.path.basename(save_dir) + ".csv")
            with open(fname, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["상품명","가격"])
                writer.writerows(self.crawled)
            self.log_msg(f"✅ CSV 저장 완료: {fname}")
            # TODO: 넘버스로 자동 열기 로직 (AppleScript or osascript)
        except Exception as e:
            self.log_msg(f"❌ CSV 저장 오류: {e}")

    def stop_crawl(self):
        self.running = False
        self.log_msg("⏸️ 크롤링 중지됨.")

    def closeEvent(self, event):
        self.driver.quit()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = CrawlerApp()
    win.show()
    sys.exit(app.exec())
