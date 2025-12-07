import sys                 # 파이썬 인터프리터/실행 환경 관련(인자, 종료 등)
import time                # 대기(sleep), 시간 제어
import re                  # 정규표현식 처리
import platform            # 운영체제 판별(macOS 여부 확인용)
import pygetwindow as gw   # OS 윈도우(창) 탐색/활성화용
import pyautogui           # 키보드/마우스 자동화(시트 입력/탭 이동/더블클릭 등)
import pyperclip           # 클립보드 복사/붙여넣기 제어
from threading import Thread  # 백그라운드 쓰레드(클릭 대기 등 비동기 처리)

import os                 # 파일/경로/환경변수
import socket             # 포트 오픈 체크(Chrome 디버그 포트)
import subprocess         # 외부 프로세스 실행(Chrome 디버그 모드 기동)
from pathlib import Path  # 경로 처리(플랫폼 독립적)
from urllib.parse import urlparse  # URL 파싱(도메인 추출 등)
from datetime import datetime      # 날짜 포맷(M/D) 작성

from PyQt6.QtGui import QKeySequence , QShortcut # PyQt6 키시퀀스(단축키 정의용)
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QTextEdit, QHBoxLayout
# ↑ PyQt6 위젯들(메인 윈도우, 버튼, 라벨, 텍스트 로그 등)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal
# ↑ Qt 코어(타이머, 시그널/슬롯 시스템)
from pynput.mouse import Listener as MouseListener
from pynput import mouse       # 마우스 이벤트(좌표, 클릭) 수신
from selenium import webdriver # 셀레니움 웹드라이버(디버그 크롬에 붙기)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# ↑ DOM 로딩/요소 대기 및 셀렉터 검색

import json                    # CDP 대상 탭 목록(/json) 파싱
import urllib.request          # CDP 엔드포인트 호출
import urllib.error            # URL 오류 처리



# =========================
# 설정값 (튜닝 포인트)
# =========================
CLICK_TIMEOUT_MS_SELECT = 5000   # [대상윈도우] 선택 클릭 타임아웃(ms)
CLICK_TIMEOUT_MS_RECORD = 10000  # [기록] 시, 시트 셀 더블클릭 유도 타임아웃(ms)

KEY_DELAY_SEC = 0.01       # 키 입력 사이 대기(IME/시트 안정성)
CLICK_STABILIZE_SEC = 0.01  # 창 활성화/포커스 후 안정 대기
NAV_DELAY_SEC = 0.005      # 탭 이동 등 네비게이션 지연

DATE_FORMAT = "M/D"        # 날짜 포맷(M/D)
FIXED_CONST_FEE = "3000"   # I열 고정 수수료 값

# 열 이동: 현재 선택 셀 기준 Tab 횟수(시트 구조에 맞춰 조정)
TABS_A_TO_B = 1
TABS_B_TO_C = 1
TABS_C_TO_F = 3
TABS_F_TO_H = 2
TABS_H_TO_I = 1
TABS_I_TO_URL = 12         # I열에서 URL 열까지 우측 12칸(요청사항 반영)

SNAP_TO_COLUMN_A = True    # (미사용 옵션 예시) A열로 스냅 이동 할 때 쓰도록 남겨둠
STRICT_REQUIRE_A = False   # (미사용 옵션 예시) A열 강제 요구 여부

# 도메인 → 라벨 매핑(시트 C열에 기록)
DOMAIN_LABELS = {
    "domeme.domeggook.com": "도매매",
    "naver.com": "네이버",
    "costco.co.kr": "코코",
    "ownerclan.com": "오너",
}


class ChromeCrawler(QWidget):          # PyQt6 메인 윈도우 클래스
    clickDetected = pyqtSignal(int, int)  # 외부 마우스 클릭 좌표를 UI 스레드로 전달하는 시그널

    def __init__(self):
        super().__init__()
        self.setWindowTitle("크롬 크롤링 도구")      # 윈도우 제목
        self.setGeometry(0, 0, 400, 500)     # 초기 위치/크기

        # -------- 상태값 --------
        self.target_title = None   # [대상윈도우] 클릭으로 얻은 OS창 제목(탭 매칭 힌트)
        self.target_window = None  # pygetwindow Window 객체
        self.driver = None         # 붙은 Selenium WebDriver 인스턴스
        self._listener = None      # 마우스 리스너(선택/기록 대기)
        self._waiting_click = False  # 클릭 대기 중 플래그
        self._click_timer = None     # 선택 타임아웃 타이머

        # -------- 크롤링 결과(시트에 쓸 값) --------
        self.crawled_title = ""   # 상품 제목
        self.crawled_price = ""   # 가격(숫자만)
        self.crawled_url = ""     # 현재 탭 URL

        # -------- UI 구성 --------
        layout = QVBoxLayout()                    # 수직 레이아웃
        self.label = QLabel("🖱 대상 윈도우: 없음")  # 현재 대상 OS창 라벨
        layout.addWidget(self.label)

        self.crawl_output = QTextEdit()           # 로그 출력 텍스트박스
        self.crawl_output.setReadOnly(True)       # 편집금지
        
        layout.addWidget(self.crawl_output)

        # 추가: Txt clear 버튼 (로그창 아래, 버튼들 위)
        self.btn_clear = QPushButton("Txt clear")
        self.btn_clear.clicked.connect(self.crawl_output.clear)  # 로그창 전체 비우기
        layout.addWidget(self.btn_clear)


        self.crawl_output.append(                 # 사용 설명 출력
            "ℹ️ 사용법:\n"
            "1) [크롬(디버그) 실행] 후 대상 페이지를 엽니다.\n"
            "2) [대상윈도우] → 대상 탭의 **본문**을 클릭(주소창/탭바 X).\n"
            "3) 크롤 직후 '대기'가 뜨면 구글시트를 **두 번 클릭**(첫 클릭 포커스, 두 번째에서 A열 입력 시작).\n"
        )

        self.btn_launch = QPushButton("크롬(디버그) 실행")
        self.btn_launch.clicked.connect(self.launch_debug_chrome)

        self.btn_test = QPushButton("기존 창 연결 테스트")
        self.btn_test.clicked.connect(self.test_attach_existing)

        # 두 버튼을 가로로 한 줄에 배치 (반반)
        row_launch = QHBoxLayout()
        row_launch.addWidget(self.btn_launch)
        row_launch.addWidget(self.btn_test)
        layout.addLayout(row_launch)


        self.btn_select = QPushButton("대상윈도우 (Shift + z)")            # 대상 OS창(크롬 탭 본문) 선택
        self.btn_select.clicked.connect(self.select_target_window)
        layout.addWidget(self.btn_select)

        self.btn_record = QPushButton("기록 (Shift + x)")                   # 시트에 기록 시작
        self.btn_record.clicked.connect(self.record_data)
        layout.addWidget(self.btn_record)

        self.btn_stop = QPushButton("STOP (프로그램 off)")                    # 앱 종료
        self.btn_stop.clicked.connect(self.close)
        layout.addWidget(self.btn_stop)

        self.setLayout(layout)                                 # 레이아웃 적용

        # ---- 단축키 바인딩 ----
        QShortcut(QKeySequence("Shift+Z"), self, activated=self.select_target_window)
        QShortcut(QKeySequence("Shift+X"), self, activated=self.record_data)


        # -------- 디버그 크롬 설정 --------
        self.DEBUGGER_ADDR = "127.0.0.1:9222"  # 디버그 어드레스(host:port)
        self.DEBUGGER_PORT = 9222              # 포트
        self.CHROME_PATHS = [                  # macOS 크롬 바이너리 후보 경로
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Google Chrome Beta.app/Contents/MacOS/Google Chrome Beta",
            "/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary",
        ]
        self.USER_DATA_DIR = str(Path("/Users/Shared/chrome_dev"))  # 별도 프로필 사용

        # -------- 선택자(사이트별/기본) --------
        self.SITE_SELECTORS = {  # 사이트별 h1/제목 후보 셀렉터
            "domeme.domeggook.com": ["#lInfoItemTitle", "h1#lInfoItemTitle", "h1"]
        }
        self.SITE_PRICE_SELECTORS = {  # 사이트별 가격 후보 셀렉터
            "domeme.domeggook.com": ["#lItemPrice", ".lItemPrice", "#lItemPriceText"]
        }
        self.DEFAULT_SELECTORS = [     # 일반적인 h1 후보 셀렉터들
            '#lInfoItemTitle', 'h1.l.infoItemTitle',
            'h1#l\\.infoItemTitle', 'h1',
            '[role="heading"][aria-level="1"]'
        ]
        self.URL_PATTERNS = [          # 특정 도메인/패턴 탭 우선 매칭
            "domeme.domeggook.com/s/", "domeme.domeggook.com"
        ]

        # 시그널 연결: 외부 마우스 클릭 → UI 스레드 핸들러
        self.clickDetected.connect(self._handle_click_on_main)

    # =============== 유틸 ===============
    def _is_macos(self) -> bool:
        return platform.system().lower() == "darwin"  # macOS 여부

    def _copy_mod(self) -> str:
        return "command" if self._is_macos() else "ctrl"  # 단축키 모디파이어 결정

    def _safe_str(self, v) -> str:
        # 예외 안전한 문자열 변환(Callable 처리/None 대비)
        try:
            if callable(v):
                v = v()
        except Exception:
            pass
        try:
            return "" if v is None else str(v)
        except Exception:
            return ""

    def _digits_only(self, s: str) -> str:
        return re.sub(r"[^0-9]", "", self._safe_str(s))  # 숫자만 추출

    def _is_int_string(self, s: str) -> bool:
        # 정수형 문자열 여부(공백/부호 허용)
        return re.fullmatch(r"\s*[+-]?\d+\s*", self._safe_str(s)) is not None

    def _today_fmt(self) -> str:
        # M/D 포맷 날짜 문자열 생성
        now = datetime.now()
        return f"{now.month}/{now.day}" if DATE_FORMAT == "M/D" else f"{now.month:02d}/{now.day:02d}"

    def _gw_get_windows_at(self, x: int, y: int):
        # 좌표(x,y)에 있는 윈도우 찾기(불가 시 활성창 폴백)
        try:
            fn = getattr(gw, "getWindowsAt", None)
            if callable(fn):
                return fn(x, y)
        except Exception:
            pass
        try:
            active = getattr(gw, "getActiveWindow", lambda: None)()
            return [active] if active else []
        except Exception:
            return []

    # ========== 안전 키보드/클립보드 (IME 문제 회피) ==========
    def _sheet_exit_edit_mode(self):
        # 시트 셀 편집모드 → 선택모드로 전환(ESC)
        pyautogui.press('esc'); time.sleep(KEY_DELAY_SEC)

    def _hotkey_mod(self, key: str):
        # ⌘/Ctrl + {key} 조합을 직접 누르는 유틸(지연 포함)
        mod = self._copy_mod()
        pyautogui.keyDown(mod); time.sleep(KEY_DELAY_SEC)
        pyautogui.press(key); time.sleep(KEY_DELAY_SEC)
        pyautogui.keyUp(mod); time.sleep(KEY_DELAY_SEC)

    def _copy_cell_to_clipboard(self):
        # 현재 선택 셀 값을 클립보드로 복사(ESC → ⌘/Ctrl+C)
        self._sheet_exit_edit_mode()
        time.sleep(KEY_DELAY_SEC)
        self._sheet_exit_edit_mode()
        time.sleep(KEY_DELAY_SEC)
        self._hotkey_mod('c')
        time.sleep(KEY_DELAY_SEC)
        self._sheet_exit_edit_mode()
        time.sleep(KEY_DELAY_SEC)

    # [중요] 시트에 값 쓰기(Enter 미사용 → 아래칸 이동 방지)
    def _set_cell_value(self, text: str):
        """
        - 편집모드 종료(ESC)
        - Delete로 기존값 삭제
        - 클립보드에 값 설정
        - 붙여넣기(⌘/Ctrl+V)만 실행(Enter 없이 확정)
        """
        self._sheet_exit_edit_mode()
        time.sleep(KEY_DELAY_SEC)
        pyperclip.copy(text or "")
        time.sleep(KEY_DELAY_SEC)
        time.sleep(KEY_DELAY_SEC)
        self._hotkey_mod('v')
        time.sleep(NAV_DELAY_SEC)
        self._sheet_exit_edit_mode()
        time.sleep(KEY_DELAY_SEC)

    def _verify_cell(self, expected: str, col_letter: str, max_retry: int = 1) -> bool:
        # 시트에서 값 재복사하여 기대값과 일치 검증(최대 재시도)
        exp_norm = (expected or "").strip()
        for attempt in range(1, max_retry + 1):
            self._copy_cell_to_clipboard()
            got_norm = (pyperclip.paste() or "").strip()
            if got_norm == exp_norm:
                print(f"{col_letter}열 : 데이터 체크 확인 (시도 {attempt}/{max_retry}) 데이터 : {exp_norm}")
                return True
            # 불일치 시 동일 값으로 다시 덮어쓰기
            self._set_cell_value(exp_norm)
            print(f"{col_letter}열 : 재시도 {attempt}/{max_retry} (읽힘='{got_norm}', 기대='{exp_norm}')")
        print(f"{col_letter}열 : 데이터 카피 오류 데이터 : {exp_norm}")
        return False

    def _press_tabs(self, n: int):
        # Tab 키 n회 눌러 우측 셀로 이동
        for _ in range(n):
            pyautogui.press('tab'); time.sleep(NAV_DELAY_SEC)

    def _go_start_of_row(self) -> bool:
        # Home(또는 macOS 폴백)으로 현재 행 A열로 이동
        try:
            pyautogui.press('home'); time.sleep(NAV_DELAY_SEC)
            if self._is_macos():
                pyautogui.hotkey('fn', 'left'); time.sleep(NAV_DELAY_SEC)
            return True
        except Exception:
            return False

    # =============== CDP(디버그 탭 탐지) ===============
    def _cdp_targets(self):
        # 디버그 포트의 /json, /json/list에서 탭 목록 조회
        urls = [f"http://{self.DEBUGGER_ADDR}/json", f"http://{self.DEBUGGER_ADDR}/json/list"]
        for u in urls:
            try:
                with urllib.request.urlopen(u, timeout=0.7) as r:
                    arr = json.loads(r.read().decode("utf-8"))
                    if isinstance(arr, list):
                        # page/background_page/other 만 사용
                        return [t for t in arr if t.get("type") in ("page", "background_page", "other")]
            except Exception:
                continue
        return []

    def _best_cdp_match(self, clicked_title: str):
        # OS창에서 얻은 제목과 CDP 목록을 비교해 가장 근접한 탭 후보 반환
        targets = self._cdp_targets()
        if not targets:
            return None
        ct = self._safe_str(clicked_title).strip()
        for t in targets:
            if ct and ct in self._safe_str(t.get("title", "")).strip():
                return t
        for t in targets:
            tt = self._safe_str(t.get("title", "")).strip()
            if tt and tt in ct:
                return t
        return targets[0]  # 그래도 없으면 첫 항목 폴백

    # =============== 디버그 크롬 ===============
    def _is_port_open(self, host: str, port: int, timeout=0.3) -> bool:
        # host:port에 TCP 연결이 되는지 체크(디버그 포트 열림 확인)
        try:
            with socket.create_connection((host, port), timeout=timeout):
                return True
        except OSError:
            return False

    def launch_debug_chrome(self):
        # 디버그 모드 크롬 실행(이미 열려 있으면 안내)
        try:
            if self._is_port_open("127.0.0.1", self.DEBUGGER_PORT):
                msg = f"ℹ️ 디버그 포트 {self.DEBUGGER_PORT}가 이미 열려 있습니다. 기존 창에 연결하세요."
                self.crawl_output.append(msg); print(msg); return

            chrome_bin = None
            for p in self.CHROME_PATHS:
                if os.path.exists(p):
                    chrome_bin = p; break
            if chrome_bin is None:
                self.crawl_output.append("⚠️ Chrome 실행 파일을 찾지 못했습니다. 경로를 확인해 주세요.")
                print("오류: Chrome 실행 파일 경로 미발견"); return

            Path(self.USER_DATA_DIR).mkdir(parents=True, exist_ok=True)  # 프로필 폴더 생성
            cmd = [chrome_bin, f"--remote-debugging-port={self.DEBUGGER_PORT}",
                   f'--user-data-dir={self.USER_DATA_DIR}', "--no-first-run", "--no-default-browser-check"]
            subprocess.Popen(  # 백그라운드 실행
                cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True
            )

            # 포트 오픈 대기(최대 ~5초)
            for _ in range(25):
                if self._is_port_open("127.0.0.1", self.DEBUGGER_PORT):
                    msg = f"✅ 디버깅 모드 Chrome 실행됨 (포트 {self.DEBUGGER_PORT})."
                    self.crawl_output.append(msg); print(msg); return
                time.sleep(0.2)

            self.crawl_output.append("⚠️ Chrome 실행은 되었지만 디버그 포트 연결을 확인하지 못했습니다.")
            print("오류: 디버그 포트 연결 확인 실패")
        except Exception as e:
            self.crawl_output.append(f"[오류] 크롬(디버그) 실행 실패: {e}")
            print(f"오류: 크롬(디버그) 실행 실패: {e}")

    def _attach_driver(self):
        # 이미 열린 디버그 포트에 WebDriver를 붙임(새 창 띄우지 않음)
        if not self._is_port_open("127.0.0.1", self.DEBUGGER_PORT):
            raise RuntimeError("디버그 포트가 열려 있지 않습니다. 먼저 '크롬(디버그) 실행'을 눌러주세요.")
        if self.driver:
            return self.driver
        options = webdriver.ChromeOptions()
        options.debugger_address = self.DEBUGGER_ADDR
        self.driver = webdriver.Chrome(options=options)
        return self.driver

    def test_attach_existing(self):
        # 현재 디버그 세션의 모든 탭 제목/URL을 로그로 출력
        try:
            driver = self._attach_driver()
            tabs_info = []
            for h in driver.window_handles:
                driver.switch_to.window(h)
                tabs_info.append(f"- {self._safe_str(driver.title).strip()}  |  {self._safe_str(driver.current_url).strip()}")
            msg = "🔗 디버그 세션 탭 목록:\n" + ("\n".join(tabs_info) if tabs_info else "(없음)")
            self.crawl_output.append(msg); print(msg)
        except Exception as e:
            self.crawl_output.append(f"[오류] 기존 창 연결 테스트 실패: {e}")
            print(f"오류: 기존 창 연결 테스트 실패: {e}")

    # =============== 대상 선택/크롤/대기 ===============
    def select_target_window(self):
        # [대상윈도우] 버튼/단축키 클릭 시 호출 → 로그 txt 창 클리어 → 클릭 대기 진입
        self.crawl_output.clear()

        # 사용자에게 "본문을 클릭"하도록 안내하고 클릭 좌표를 기다림
        self.label.setText("🔍 본문을 클릭하세요 (주소창 X). 5초 내 미클릭 시 경고.")
        self.crawl_output.append("🖱 **크롤링할 크롬 탭의 본문**을 클릭해 주세요. (5초 내)\n")
        self.showMinimized()        # 앱 최소화(클릭 방해 줄이기)
        self._waiting_click = True  # 클릭 대기 시작

        # 타임아웃 타이머 준비
        if self._click_timer is None:
            self._click_timer = QTimer(self)
            self._click_timer.setSingleShot(True)
            self._click_timer.timeout.connect(self._on_click_timeout_select)
        self._click_timer.start(CLICK_TIMEOUT_MS_SELECT)

        # pynput로 전역 마우스 클릭 이벤트 후킹
        def on_click(x, y, button, pressed):
            if pressed and self._waiting_click:
                self.clickDetected.emit(int(x), int(y))  # 좌표를 시그널로 전송
        self._listener = MouseListener(on_click=on_click)
        self._listener.start()

    def _on_click_timeout_select(self):
        # [대상윈도우] 클릭 대기 타임아웃 시 호출
        if not self._waiting_click:
            return
        self._waiting_click = False
        try:
            if self._listener: self._listener.stop()
        except Exception:
            pass
        finally:
            self._listener = None
        self.crawl_output.append("⏰ 5초 내 클릭이 감지되지 않았습니다. 다시 [대상윈도우]를 눌러 본문을 클릭하세요.")
        print("⏰ (타임아웃) 대상윈도우 클릭 미감지")

    def _handle_click_on_main(self, x: int, y: int):
        # 전역 클릭 좌표 수신 → 해당 위치의 OS창/제목을 기반으로 대상 설정
        if not self._waiting_click:
            return
        self._waiting_click = False
        if self._click_timer and self._click_timer.isActive():
            self._click_timer.stop()
        try:
            if self._listener: self._listener.stop()
        except Exception:
            pass
        finally:
            self._listener = None

        wins_at = self._gw_get_windows_at(x, y)     # 좌표에 있는 윈도우 목록
        win = wins_at[0] if wins_at else None       # 첫 윈도우 선택
        picked_title = self._safe_str(getattr(win, "title", "")) if win else ""
        if not picked_title:
            # 본문이 아닌 주소창/탭바를 클릭했거나 권한 문제일 수 있음
            self.crawl_output.append(
                "❌ 클릭 지점에서 활성 창 제목을 찾지 못했습니다.\n"
                " - 본문을 클릭했는지 확인하세요.\n"
                " - (macOS) 손쉬운 사용/입력 모니터링 권한 확인."
            )
            print("❌ 클릭 창 탐지 실패: 제목 없음"); return

        self.target_window = win        # 대상 OS창 보관
        self.target_title = picked_title
        self.label.setText(f"🎯 대상 윈도우: {self.target_title}")

        self.showNormal(); self.raise_(); self.activateWindow()  # 앱 전면 복귀
        self.crawl_data()                 # 즉시 크롤
        self._start_record_wait(auto_trigger=True)  # 시트 기록 대기 진입

    # =============== 셀렉터 ===============
    def _selectors_for_url(self, url):
        # URL 도메인에 맞는 사이트 전용 선택자 + 기본 선택자 병합(중복 제거)
        host = urlparse(url).netloc if url else ""
        site_specific = []
        for key, sels in self.SITE_SELECTORS.items():
            if key in host:
                site_specific += sels
        seen, ordered = set(), []
        for sel in site_specific + self.DEFAULT_SELECTORS:
            if sel not in seen:
                seen.add(sel); ordered.append(sel)
        return ordered

    def _price_selectors_for_url(self, url):
        # 가격용 선택자(사이트 전용 + 일반 후보)
        host = urlparse(url).netloc if url else ""
        site_specific = []
        for key, sels in self.SITE_PRICE_SELECTORS.items():
            if key in host:
                site_specific += sels
        general = ["#lItemPrice", ".lItemPrice", ".price .num", ".price-value", ".final_price",
                   ".sale_price", ".price", "[data-testid='price']"]
        seen, ordered = set(), []
        for sel in site_specific + general:
            if sel not in seen:
                seen.add(sel); ordered.append(sel)
        return ordered

    # ===== 클릭좌표 창 활성화/포커스 보정 =====
    def _activate_window_at(self, x: int, y: int):
        """
        - 좌표의 OS 창 활성화
        - 해당 지점 실제 클릭으로 시트 그리드 포커스 강제
        """
        try:
            wins = self._gw_get_windows_at(x, y)
            w = wins[0] if wins else None
            if w:
                w.activate()
                time.sleep(0.2)
            pyautogui.click(x, y)   # 포커스 강제
            time.sleep(0.15)
            try:
                active = getattr(gw, "getActiveWindow", lambda: None)()
                print(f"[DEBUG] 활성창: {getattr(active, 'title', '(제목없음)')}")
            except Exception:
                pass
        except Exception as e:
            print(f"[경고] 창 활성화/포커스 실패: {e}")

    # =============== 크롤 ===============
    def crawl_data(self):
        # 선택된 OS창 제목/창 정보 기반으로 CDP 탭 매칭 → 제목/가격/URL 추출
        if not self.target_window and not self.target_title:
            return
        try:
            if self.target_window:
                try:
                    self.target_window.activate(); time.sleep(0.2)
                except Exception:
                    pass

            driver = self._attach_driver()  # 디버그 포트의 WebDriver 확보

            # 탭 매칭: URL 패턴 → 창제목 포함 → CDP(/json)
            self.crawl_output.append("🧭 탭 매칭: URL패턴 → 제목 포함 → CDP(/json) 순")
            end_time = time.time() + 5.0
            target_handle = None

            # 1) URL 패턴 우선
            if self.URL_PATTERNS:
                while time.time() < end_time and not target_handle:
                    for h in driver.window_handles:
                        driver.switch_to.window(h)
                        if any(p in (driver.current_url or "") for p in self.URL_PATTERNS):
                            target_handle = h; break
                    if not target_handle:
                        time.sleep(0.2)

            # 2) 창 제목 포함 매칭
            if not target_handle:
                end_time2 = time.time() + 5.0
                want = self._safe_str(self.target_title).strip()
                while time.time() < end_time2 and not target_handle:
                    for h in driver.window_handles:
                        driver.switch_to.window(h)
                        if want and want in self._safe_str(driver.title).strip():
                            target_handle = h; break
                    if not target_handle:
                        time.sleep(0.2)

            # 3) CDP 목록 기반 URL 정합
            if not target_handle:
                cdp = self._best_cdp_match(self.target_title)
                if cdp:
                    tu = self._safe_str(cdp.get("url")).strip()
                    for h in driver.window_handles:
                        driver.switch_to.window(h)
                        cur = self._safe_str(driver.current_url).strip()
                        if cur == tu or (tu and cur.startswith(tu.split("#")[0])):
                            target_handle = h; break

            if not target_handle:
                self.crawl_output.append("❌ 5초 내 '대상 탭'을 찾지 못했습니다.")
                return

            driver.switch_to.window(target_handle)  # 대상 탭 활성화

            current_url = self._safe_str(driver.current_url).strip()
            self.crawled_url = current_url
            self.crawl_output.append(f"🔗 URL: {current_url}")

            # 접근 제한 페이지 제외(chrome://, pdf 등)
            blocked = ("chrome://", "chrome-extension://", "edge://", "about:", "data:")
            if any(current_url.startswith(s) for s in blocked) or current_url.lower().endswith(".pdf"):
                self.crawl_output.append("❌ 이 페이지는 DOM 접근이 제한됩니다.")
                return

            # DOM 로딩 상태 안정화 대기
            try:
                WebDriverWait(driver, 3).until(
                    lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
                )
            except Exception:
                pass

            # 제목 추출: 사이트별 → 기본 → <h1>
            title_value = ""
            wait = WebDriverWait(driver, 5)
            for sel in self._selectors_for_url(current_url):
                try:
                    el = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, sel)))
                    title_value = (el.text or "").strip()
                    if title_value:
                        break
                except Exception:
                    continue
            if not title_value:
                try:
                    el = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.TAG_NAME, "h1")))
                    title_value = (el.text or "").strip()
                except Exception:
                    title_value = ""
            self.crawled_title = title_value
            self.crawl_output.append(f"🟢 제목: {self.crawled_title or '(없음)'}")

            # 가격 추출: 사이트별 → 일반 후보 → 본문 텍스트 Regex(원/₩)
            price_digits = ""
            wait_p = WebDriverWait(driver, 3)
            for sel in self._price_selectors_for_url(current_url):
                try:
                    el = wait_p.until(EC.visibility_of_element_located((By.CSS_SELECTOR, sel)))
                    txt = (el.text or "").strip()
                    if not txt:
                        # innerText 보조
                        txt = (driver.execute_script(
                            "const e=document.querySelector(arguments[0]); return e?(e.innerText||e.textContent||''):'';", sel
                        ) or "").strip()
                    if txt:
                        price_digits = re.sub(r"[^0-9]", "", txt)
                        if price_digits:
                            break
                except Exception:
                    continue
            if not price_digits:
                try:
                    body = (driver.execute_script(
                        "return (document.body && document.body.innerText) ? document.body.innerText : '';"
                    ) or "")
                    m = re.search(r'([0-9]{1,3}(?:,[0-9]{3})+|[0-9]+)\s*원', body)
                    if not m:
                        m = re.search(r'₩\s*([0-9]{1,3}(?:,[0-9]{3})+|[0-9]+)', body)
                    if m:
                        price_digits = re.sub(r"[^0-9]", "", m.group(1))
                except Exception:
                    pass
            self.crawled_price = price_digits
            self.crawl_output.append(f"💰 가격(숫자만): {self.crawled_price or '(없음)'}")

            # 요약 블록 출력
            self.crawl_output.append("—" * 40)
            self.crawl_output.append(f"제목: {self.crawled_title or '(없음)'}")
            self.crawl_output.append(f"가격(숫자만): {self.crawled_price or '(없음)'}")
            self.crawl_output.append(f"URL: {self.crawled_url or '(없음)'}")
            self.crawl_output.append("—" * 40)

        except Exception as e:
            self.crawl_output.append(f"[오류] 크롤링 실패: {e}")

    # =============== 자동/수동 기록(두 번째 클릭에서 시작) ===============
    def _start_record_wait(self, auto_trigger: bool):
        # 사용자가 시트에서 시작 셀을 두 번 클릭하도록 유도(10초 타임아웃)
        self.crawl_output.append("⏳ 대기: 구글시트에서 **기록 시작 셀**을 클릭해 주세요. (10초 내)\n")
        print("⌛ [DEBUG] 구글시트 셀 클릭 대기 시작 (첫 클릭은 포커스, 두 번째 클릭에서 A열 입력 시작)")

        today_str = self._today_fmt()  # B열에 쓸 날짜 문자열

        # 실제 채우기 시퀀스(두 번째 클릭 좌표를 받아 포커스 보정 후 입력)
        def do_fill_sequence(start_x: int, start_y: int):
            try:
                # 좌표 창 활성화 + 포커스 강제
                self._activate_window_at(start_x, start_y)
                time.sleep(CLICK_STABILIZE_SEC)

                # === A열: 위 셀 값 정수 +1 ===
                print("A 열. 시작")
                pyautogui.press('up'); time.sleep(NAV_DELAY_SEC)
                self._sheet_exit_edit_mode()
                time.sleep(KEY_DELAY_SEC)
                self._sheet_exit_edit_mode()
                self._copy_cell_to_clipboard(); time.sleep(NAV_DELAY_SEC)
                above_val_raw = pyperclip.paste()
                ab= int(above_val_raw)
                print('읽은 수', ab)
                detected_int = int(above_val_raw) if self._is_int_string(above_val_raw) else 0
                print(f"[Sheets] 위 셀 정수 감지: {detected_int}")
                pyautogui.press('down'); time.sleep(NAV_DELAY_SEC)
                self._sheet_exit_edit_mode()
                time.sleep(KEY_DELAY_SEC)
                a_value = str(detected_int + 1)
                self._set_cell_value(a_value)
                time.sleep(NAV_DELAY_SEC)
                self._verify_cell(a_value, "A")
                time.sleep(NAV_DELAY_SEC)


                # === B열: 날짜 ===
                self._press_tabs(TABS_A_TO_B)
                self._sheet_exit_edit_mode()
                time.sleep(KEY_DELAY_SEC)
                self._set_cell_value(today_str)
                time.sleep(NAV_DELAY_SEC)
                self._verify_cell(today_str, "B")
                time.sleep(NAV_DELAY_SEC)
                
                # === C열: 도메인 라벨 ===
                self._press_tabs(TABS_B_TO_C)
                host = urlparse(self.crawled_url or "").netloc.lower()
                label = ""
                for dom, lab in DOMAIN_LABELS.items():
                    if dom in host:
                        label = lab; break
                self._sheet_exit_edit_mode()
                time.sleep(KEY_DELAY_SEC)
                self._set_cell_value(label)
                time.sleep(NAV_DELAY_SEC)
                self._verify_cell(label, "C")
                time.sleep(NAV_DELAY_SEC)
                
                # === F열: 제목 ===
                self._press_tabs(TABS_C_TO_F)
                self._sheet_exit_edit_mode()
                time.sleep(KEY_DELAY_SEC)
                self._set_cell_value(self.crawled_title or "")
                time.sleep(NAV_DELAY_SEC)
                self._verify_cell(self.crawled_title or "", "F")
                time.sleep(NAV_DELAY_SEC)
                self._sheet_exit_edit_mode()
                time.sleep(KEY_DELAY_SEC)

                # === H열: 가격 ===
                self._press_tabs(TABS_F_TO_H)
                self._sheet_exit_edit_mode()
                time.sleep(KEY_DELAY_SEC)
                self._set_cell_value(self.crawled_price or "")
                time.sleep(NAV_DELAY_SEC)
                self._verify_cell(self.crawled_price or "", "H")
                time.sleep(NAV_DELAY_SEC)
 
                # === I열: 고정비 3000 ===
                self._press_tabs(TABS_H_TO_I)
                self._sheet_exit_edit_mode()
                time.sleep(KEY_DELAY_SEC)
                self._set_cell_value(FIXED_CONST_FEE)
                time.sleep(NAV_DELAY_SEC)
                self._verify_cell(FIXED_CONST_FEE, "I")
                time.sleep(NAV_DELAY_SEC)

                # === URL 열: I에서 +12칸 이동 ===
                self._press_tabs(TABS_I_TO_URL)
                self._set_cell_value(self.crawled_url or "")
                time.sleep(NAV_DELAY_SEC)
                self._verify_cell(self.crawled_url or "", "URL-열")

                self.crawl_output.append("✅ 구글시트 자동 입력 완료.")
                print("✅ [DEBUG] 구글시트 자동 입력 완료")

            except Exception as e:
                self.crawl_output.append(f"[오류] 구글시트 자동 입력 실패: {e}")
                print(f"[오류] 구글시트 자동 입력 실패: {e}")

        # 타임아웃 플래그(쓰레드 종료 여부 통지용)
        timed_out = {"v": False}

        def wait_for_click_then_fill():
            # 전역 클릭 이벤트를 순차 수신 → 두 번째 클릭에서 do_fill_sequence 실행
            start_ts = time.time()
            click_count = 0
            last_x, last_y = 0, 0
            with mouse.Events() as events:
                for event in events:
                    if (time.time() - start_ts) * 1000 >= CLICK_TIMEOUT_MS_RECORD:
                        timed_out["v"] = True
                        break
                    if isinstance(event, mouse.Events.Click) and event.pressed:
                        click_count += 1
                        last_x, last_y = int(event.x), int(event.y)
                        if click_count == 1:
                            print("[Sheets] 첫 클릭 감지 → 시트 창 포커스됨. 이제 **A열의 셀을 한 번 더 클릭**하세요.")
                            continue
                        do_fill_sequence(last_x, last_y)
                        return

        # 클릭 대기를 별도 쓰레드로 수행(메인 UI 블로킹 방지)
        t = Thread(target=wait_for_click_then_fill, daemon=True)
        t.start()

        # 타임아웃 폴링: 쓰레드 상태를 주기적으로 확인
        def poll_timeout():
            if not t.is_alive():
                if timed_out["v"]:
                    self.crawl_output.append("⏰ (자동 대기) 10초 내 **두 번째 클릭**이 감지되지 않았습니다. [기록] 버튼으로 재시도하세요.")
                    print("⏰ (타임아웃) 두 번째 클릭 미감지")
                return
            QTimer.singleShot(300, poll_timeout)
        QTimer.singleShot(300, poll_timeout)

    def record_data(self):
        # [기록] 버튼: 크롤된 최소 데이터(제목/URL)가 있어야 진행
        if not (self.crawled_title and self.crawled_url):
            self.crawl_output.append("⚠️ 먼저 [대상윈도우]로 제목/가격/URL을 크롤링해 주세요.")
            return
        self._start_record_wait(auto_trigger=False)  # 수동 기록 대기 진입


# =============== 엔트리 포인트 ===============
if __name__ == "__main__":
    app = QApplication(sys.argv)  # Qt 앱 생성
    win = ChromeCrawler()         # 메인 윈도우 인스턴스
    win.show()                    # 창 표시
    sys.exit(app.exec())          # 이벤트 루프 진입 후 종료 코드 반환
