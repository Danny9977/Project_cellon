from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
import matplotlib.pyplot as plt

# 크롬 드라이버 경로 필요: chromedriver를 환경에 맞게 설치하고 경로 수정
driver = webdriver.Chrome()

try:
    # 1) KB 부동산 매매지수 페이지 접속
    # url = "https://nland.kbstar.com/quics?page=C021342"
    url = "https://data.kbland.kr/kbstats/wmh?tIdx=HT01&tsIdx=weekAptSalePriceInx"
    driver.get(url)
    time.sleep(10)  # 페이지 로드 대기

    # 2) 데이터 테이블 탐색 (※ 실제 페이지에서 확인한 selector 기반 예시)
    # 매매지수가 들어간 테이블 선택 (사이트 구조에 따라 변경 필요)
    rows = driver.find_elements(By.CSS_SELECTOR, "table.table tbody tr")
    if not rows:
        print("테이블 데이터를 찾지 못했습니다. 페이지 구조가 변경되었을 수 있습니다.")
        # driver.quit()
        # exit()

    dates, indices = [], []
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) >= 2:
            date_text = cells[0].text.strip()
            idx_text = cells[1].text.strip().replace(',', '').replace('-', '')
            if date_text and idx_text.isdigit():
                dates.append(date_text)
                indices.append(float(idx_text))

    # 3) 데이터 정리
    df = pd.DataFrame({"날짜": dates, "매매지수": indices})
    df["날짜"] = pd.to_datetime(df["날짜"], errors='coerce')  # 날짜 변환
    df = df.dropna().sort_values(by="날짜")  # 날짜 누락/오류 제거

    print(df)

    # 4) 그래프 그리기
    plt.figure(figsize=(12,6))
    plt.plot(df["날짜"], df["매매지수"], marker='o')
    plt.title("KB 부동산 매매지수 추이")
    plt.xlabel("날짜")
    plt.ylabel("매매지수")
    plt.grid(True)
    plt.tight_layout()
    plt.show()

finally:
    driver.quit()
