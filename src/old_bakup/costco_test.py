# 코스트코 이미지 캡쳐하는 코드 - 잘됨

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from PIL import Image
import base64
import io
import os

# 크롬 드라이버 경로 설정
chrome_driver_path = '/usr/local/bin/chromedriver'  # 본인 경로에 맞게 수정
url = 'https://www.costco.co.kr/p/387546'
save_path = '/Users/jeehoonkim/Desktop/Python_Project/product_image.png'

# 브라우저 옵션 설정
options = webdriver.ChromeOptions()
options.add_argument("--headless=new")  # 창 없이 실행 (신버전 방식)
options.add_argument("--window-size=1920,3000")  # 충분히 긴 화면으로 전체 캡처 유도
options.add_argument("--force-device-scale-factor=1")  # Mac retina 스케일 보정

driver = webdriver.Chrome(service=Service(chrome_driver_path), options=options)
driver.get(url)
time.sleep(5)

try:
    # 대표 이미지 요소 찾기
    img_element = driver.find_element(By.XPATH, '/html/body/main/div[4]/sip-product-details-page/sip-product-details/div/sip-product-image-panel/div/div/div[1]/div[1]/div[1]/div/sip-image-zoom/div/sip-media[1]/picture/img')

    # 스크롤 이동
    driver.execute_script("arguments[0].scrollIntoView(true);", img_element)
    time.sleep(1)

    # 요소 위치 및 크기 측정
    location = img_element.location_once_scrolled_into_view
    size = img_element.size

    # Chrome DevTools Protocol을 통해 전체 스크린샷 (base64)
    screenshot_base64 = driver.get_screenshot_as_base64()
    screenshot_data = base64.b64decode(screenshot_base64)

    # 이미지 열기
    image = Image.open(io.BytesIO(screenshot_data))

    # 정확한 크기로 자르기
    left = int(location['x'])
    top = int(location['y'])
    right = left + int(size['width'])
    bottom = top + int(size['height'])

    cropped = image.crop((left, top, right, bottom))
    cropped.save(save_path)

    print(f"✅ 이미지 저장 성공: {save_path}")

except Exception as e:
    print(f"❌ 오류 발생: {e}")

finally:
    driver.quit()
