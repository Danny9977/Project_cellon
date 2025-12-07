from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time

# 크롬 드라이버 경로를 본인 환경에 맞게 수정하세요.
chrome_driver_path = "/usr/local/bin/chromedriver"

options = Options()
options.add_argument("--start-maximized")  # 필요하면 창을 최대화해서 보세요.

driver = webdriver.Chrome(service=Service(chrome_driver_path), options=options)

try:
    # 페이지 접속
    url = "https://data.kbland.kr/kbstats/wmh?tIdx=HT01&tsIdx=weekAptSalePriceInx"
    driver.get(url)

    # "서울" 요소가 로드될 때까지 대기
    wait = WebDriverWait(driver, 15)
    seoul_element = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//div[contains(@class, 'bodyitem') and contains(., '서울')]")
        )
    )

    # "서울" 클릭 (seoul_element)
    button_text = seoul_element.text.strip()  # 버튼에 적힌 글자 추출 -> 캡쳐 파일명을 위함
    seoul_element.click()
    print("서울 클릭 완료!")

    # 차트 부분 로드 대기 
    chartin = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.chartin")))
    
    time.sleep(2)  # 차트가 완전히 로드될 시간 추가
    
     # 파일명 생성: '버튼명_YYYYMMDD.jpeg'
    today_str = datetime.now().strftime("%Y%m%d")
    filename = f"{button_text}_{today_str}.jpeg"

    # chartbox 영역 스크린샷 찍기
    chartin.screenshot(filename)
    print(f"차트 스크린샷 저장 완료: {filename}")

except Exception as e:
    print(f"오류 발생: {e}")

finally:
    # 필요하면 드라이버 종료 전에 sleep으로 확인시간 추가
    # import time; time.sleep(5)
    # driver.quit()
    print(f"성공 성공 성공")
