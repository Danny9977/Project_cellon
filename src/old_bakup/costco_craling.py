# -*- coding: utf-8 -*-

import sys
import os
from datetime import datetime
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QTextEdit
from PyQt6.QtCore import QThread, pyqtSignal
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
import requests
import csv
from rembg import remove
from io import BytesIO
import subprocess

# 크롤링 쓰레드 클래스
class CrawlerThread(QThread):
    data_collected = pyqtSignal(str)

    def __init__(self, url):
        super().__init__()
        self.url = url
        self._is_running = True

    def run(self):
        try:
            options = Options()
            options.add_argument('--headless=new')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            service = Service('/usr/local/bin/chromedriver')  # User 코드 삽입 필요
            driver = webdriver.Chrome(service=service, options=options)
            driver.set_window_size(1920, 1080)
            driver.get(self.url)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'h1.product-name'))
            )

            name = driver.find_element(By.CSS_SELECTOR, 'h1.product-name').text
            price = driver.find_element(By.CSS_SELECTOR, 'span.notranslate.ng-star-inserted').text
            img_elem = driver.find_element(By.CSS_SELECTOR, 'img.ng-star-inserted')
            img_url = img_elem.get_attribute('src')


            print("엘레먼트 찾기 완료")

            now = datetime.now().strftime('%Y%m%d')
            folder_path = f'/Users/jeehoonkim/Desktop/Python_Project/코스트코상품_{now}'
            os.makedirs(folder_path, exist_ok=True)

            # 이미지 다운로드 및 누끼 처리
            try:
                img_response = requests.get(img_url, timeout=10)
                input_img = Image.open(BytesIO(img_response.content))
                output_img = remove(input_img)
                final_img = output_img.convert("RGB")
                final_img.save(os.path.join(folder_path, f'{name}.jpg'), 'JPEG')
            except Exception as e:
                print('이미지 처리 에러:', e)
            print("이미지 저장 완료")

            # 상품 상세정보 캡처
            try:
                detail_elem = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'wrapper_itemDes')))
                detail_elem.screenshot(os.path.join(folder_path, f'{name}_상세정보.jpg'))
            except Exception as e:
                print('상품상세정보 캡처 실패:', e)
            print("상품 상세정보 완료")

            # 스펙 캡처
            try:
                spec_elem = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'product-classification-wrapper')))
                spec_elem.screenshot(os.path.join(folder_path, f'{name}_스펙.jpg'))
            except Exception as e:
                print('스펙 캡처 실패:', e)
            print("스펙 완료")

            # 데이터 CSV 저장
            csv_path = os.path.join(folder_path, f'코스트코상품_{now}.csv')
            with open(csv_path, 'a', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow([name, price, self.url])

            self.data_collected.emit(f'{name} | {price} | {self.url}')
            driver.quit()
        except Exception as e:
            print('에러 발생:', e)
        print("파일저장 완료")

    def stop(self):
        self._is_running = False
        self.terminate()


class CostcoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.crawler_thread = None
        self.crawled_data = []

    def initUI(self):
        self.setWindowTitle('코스트코 크롤링')
        layout = QVBoxLayout()

        self.url_input = QLineEdit(self)
        self.url_input.setPlaceholderText('크롤할 상품 URL 입력')
        layout.addWidget(self.url_input)

        self.crawl_btn = QPushButton('크롤링 시작', self)
        self.crawl_btn.clicked.connect(self.start_crawling)
        layout.addWidget(self.crawl_btn)

        self.stop_btn = QPushButton('크롤링 스탑', self)
        self.stop_btn.clicked.connect(self.stop_crawling)
        layout.addWidget(self.stop_btn)

        self.export_btn = QPushButton('data 보내기', self)
        self.export_btn.clicked.connect(self.export_to_numbers)
        layout.addWidget(self.export_btn)

        self.result_box = QTextEdit(self)
        self.result_box.setReadOnly(True)
        layout.addWidget(self.result_box)

        self.setLayout(layout)

    def start_crawling(self):
        url = self.url_input.text()
        if url:
            self.crawler_thread = CrawlerThread(url)
            self.crawler_thread.data_collected.connect(self.display_result)
            self.crawler_thread.start()

    def stop_crawling(self):
        if self.crawler_thread:
            self.crawler_thread.stop()

    def display_result(self, data):
        self.result_box.append(data)
        self.crawled_data.append(data.split(' | '))

    def export_to_numbers(self):
        now = datetime.now().strftime('%Y%m%d')
        folder_path = f'/Users/jeehoonkim/Desktop/Python_Project/코스트코상품_{now}'
        os.makedirs(folder_path, exist_ok=True)
        csv_path = os.path.join(folder_path, f'코스트코상품_{now}.csv')
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['상품명', '가격', 'URL'])
            for row in self.crawled_data:
                writer.writerow(row)

        try:
            subprocess.run(["open", csv_path], check=True)
        except Exception as e:
            print("Numbers 파일 열기 실패:", e)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = CostcoApp()
    window.show()
    sys.exit(app.exec())
