import sys
import os
import re
import time
import json
import platform
import socket
import subprocess
from pathlib import Path
from urllib.parse import urlparse  # <- 유지
from datetime import datetime

# ==== PyQt6 ====
from PyQt6.QtGui import QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QTextEdit, QHBoxLayout
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal

# ==== UI/OS/입력 ====
import pygetwindow as gw
import pyautogui
import pyperclip
from pynput import mouse
from pynput.mouse import Listener as MouseListener

# ==== Selenium ====
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==== Google Sheets ====
import gspread
from google.oauth2.service_account import Credentials
from google.auth.exceptions import TransportError

# ==== 쿠팡 OpenAPI ====
import requests
import hmac, hashlib, base64
from urllib.parse import urlencode, quote  # canonical query 생성을 위해 quote 사용

# =========================
# 설정값 (튜닝 포인트)
# =========================
# --- Google Sheets ---
SERVICE_ACCOUNT_JSON = "/Users/jeehoonkim/Desktop/Python_Project/api/google_api/service_account.json"  # 서비스계정 키 경로
SHEET_ID = "1OEg01RdJyesSy7iQSEyQHdYpCX5MSsNUfD0lkUYq8CM"  # 스프레드시트 ID
WORKSHEET_NAME = "소싱상품목록"  # 시트 탭 이름

# --- 크롬 디버그 포트/경로 ---
DEBUGGER_ADDR = "127.0.0.1:9222"
DEBUGGER_PORT = 9222
CHROME_PATHS = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Google Chrome Beta.app/Contents/MacOS/Google Chrome Beta",
    "/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary",
]
USER_DATA_DIR = str(Path("/Users/Shared/chrome_dev"))

# --- 지연/타임아웃 ---
CLICK_TIMEOUT_MS_SELECT = 5000   # 대상윈도우 선택(본문 클릭) 대기 타임아웃
CLICK_TIMEOUT_MS_RECORD = 10000  # 시트 클릭 대기 타임아웃
KEY_DELAY_SEC = 0.01
CLICK_STABILIZE_SEC = 0.01
NAV_DELAY_SEC = 0.005

DATE_FORMAT = "M/D"        # 날짜 포맷
FIXED_CONST_FEE = "3000"   # I열 고정 수수료

# --- URL→라벨 매핑(C열) ---
DOMAIN_LABELS = {
    "domeme.domeggook.com": "도매매",
    "naver.com": "네이버",
    "costco.co.kr": "코코",
    "ownerclan.com": "오너",
}

# --- 크롤링용 기본/사이트별 셀렉터 ---
SITE_SELECTORS = {
    "domeme.domeggook.com": ["#lInfoItemTitle", "h1#lInfoItemTitle", "h1"]
}
SITE_PRICE_SELECTORS = {
    "domeme.domeggook.com": ["#lItemPrice", ".lItemPrice", "#lItemPriceText"]
}
DEFAULT_SELECTORS = [
    '#lInfoItemTitle', 'h1.l.infoItemTitle',
    'h1#l\\.infoItemTitle', 'h1',
    '[role="heading"][aria-level="1"]'
]
URL_PATTERNS = ["domeme.domeggook.com/s/", "domeme.domeggook.com"]

# --- Coupang Open API (Wing) ---
COUPANG_BASE_URL = "https://api-gateway.coupang.com"
COUPANG_KEYS_JSON = str(Path("/Users/jeehoonkim/Desktop/api/coupang_api/coupang_keys.json"))

try:
    with open(COUPANG_KEYS_JSON, "r", encoding="utf-8") as f:
        coupang_keys = json.load(f)
        COUPANG_VENDOR_ID = (coupang_keys.get("vendor_id") or "").strip()
        COUPANG_ACCESS_KEY = (coupang_keys.get("access_key") or "").strip()
        COUPANG_SECRET_KEY = (coupang_keys.get("secret_key") or "").strip()
except Exception as e:
    print(f"❌ 쿠팡 키 파일을 불러오지 못했습니다: {e}")
    COUPANG_VENDOR_ID = COUPANG_ACCESS_KEY = COUPANG_SECRET_KEY = None

COUPANG_LOOKBACK_DAYS = 7
CP_STATUS_MAP = {
    "ACCEPT": "상품준비중",
    "INSTRUCT": "배송지시",
    "DELIVERING": "배송중",
    "DELIVERED": "배송완료",
}
CP_QUERY_STATUSES = ["ACCEPT", "INSTRUCT", "DELIVERING", "DELIVERED"]
COUPANG_WS_NAME = "쿠팡주문현황"

# =========================
# 유틸 함수
# =========================
def is_macos() -> bool:
    return platform.system().lower() == "darwin"

def safe_str(v) -> str:
    try:
        if callable(v): v = v()
    except Exception:
        pass
    try:
        return "" if v is None else str(v)
    except Exception:
        return ""

def digits_only(s: str) -> str:
    return re.sub(r"[^0-9]", "", safe_str(s))

def is_int_string(s: str) -> bool:
    return re.fullmatch(r"\s*[+-]?\d+\s*", safe_str(s)) is not None

def today_fmt() -> str:
    now = datetime.now()
    return f"{now.month}/{now.day}" if DATE_FORMAT == "M/D" else f"{now.month:02d}/{now.day:02d}"

def is_port_open(host: str, port: int, timeout=0.3) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False

def selectors_for_url(url: str):
    host = urlparse(url).netloc if url else ""
    site_specific = []
    for key, sels in SITE_SELECTORS.items():
        if key in host:
            site_specific += sels
    seen, ordered = set(), []
    for sel in site_specific + DEFAULT_SELECTORS:
        if sel not in seen:
            seen.add(sel); ordered.append(sel)
    return ordered

def price_selectors_for_url(url: str):
    host = urlparse(url).netloc if url else ""
    site_specific = []
    for key, sels in SITE_PRICE_SELECTORS.items():
        if key in host:
            site_specific += sels
    general = ["#lItemPrice", ".lItemPrice", ".price .num", ".price-value", ".final_price",
               ".sale_price", ".price", "[data-testid='price']"]
    seen, ordered = set(), []
    for sel in site_specific + general:
        if sel not in seen:
            seen.add(sel); ordered.append(sel)
    return ordered

def label_for_domain(url: str) -> str:
    host = urlparse(url or "").netloc.lower()
    for dom, lab in DOMAIN_LABELS.items():
        if dom in host:
            return lab
    return ""

def _mask(s: str, left: int = 4, right: int = 3) -> str:
    """키 마스킹: 앞/뒤 일부만 보이고 나머지는 * 처리"""
    s = str(s or "")
    if len(s) <= left + right:
        return "*" * len(s)
    return s[:left] + "*" * (len(s) - left - right) + s[-right:]

# === 쿠팡 요청 / 호출 헬퍼 =====
# RFC3986 인코딩 + CEA 개행 포맷 서명 (가이드 고정형)

def _cp_rfc3986_encode(s: str) -> str:
    """RFC3986 스타일 인코딩: 공백→%20, '+' 금지, safe='-_.~'"""
    return quote(str(s), safe='-_.~')

def _cp_canonical_query_for_signature(params: dict | None) -> str:
    """
    쿠팡 서명/URL용 canonical query:
    - (key, value) 모두 문자열화
    - key → value 순으로 정렬
    - RFC3986 인코딩 (safe='-_.~'), 공백은 %20
    - 앞에 '?'를 붙이지 않음
    """
    if not params:
        return ""
    items = []
    for k, v in params.items():
        if isinstance(v, (list, tuple)):
            for vv in v:
                items.append((str(k), str(vv)))
        else:
            items.append((str(k), str(v)))
    # key, value 기준 안정 정렬
    items.sort(key=lambda kv: (kv[0], kv[1]))
    return "&".join(f"{_cp_rfc3986_encode(k)}={_cp_rfc3986_encode(v)}" for k, v in items)

from datetime import datetime, timezone

def _cp_signed_headers(method: str, path: str, params_or_query: dict | str | None,
                       access_key: str, secret_key: str, *,
                       vendor_id: str | None = None) -> tuple[dict, str, str]:
    """
    Coupang CEA 서명 헤더 생성(가이드 포맷 고정):
      message = "{METHOD}\n{PATH}\n{QUERY}\n{SIGNED_DATE}"
        - QUERY: '?' 없이 canonical query (RFC3986, 정렬). 문자열을 넘긴 경우 그대로 사용.
        - SIGNED_DATE: UTC, "YYYY-MM-DDTHH:MM:SSZ"
    반환: (headers, signed_date, canonical_query)
    """
    signed_date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    if isinstance(params_or_query, dict):
        canonical_query = _cp_canonical_query_for_signature(params_or_query)
    elif isinstance(params_or_query, str):
        canonical_query = params_or_query
    else:
        canonical_query = ""

    msg = f"{method.upper()}\n{path}\n{canonical_query}\n{signed_date}"

    signature = base64.b64encode(
        hmac.new(secret_key.encode("utf-8"), msg.encode("utf-8"), hashlib.sha256).digest()
    ).decode("utf-8")

    headers = {
        "Authorization": (
            f"CEA algorithm=HmacSHA256, access-key={access_key}, "
            f"signed-date={signed_date}, signature={signature}"
        ),
        "x-coupang-date": signed_date,  # 게이트웨이에 따라 사실상 필수처럼 동작
        "Content-Type": "application/json;charset=UTF-8",
    }
    if vendor_id:
        headers["X-Requested-By"] = vendor_id  # 일부 엔드포인트에서 요구됨

    return headers, signed_date, canonical_query

def _cp_request(method: str, path: str, params: dict | None) -> dict:
    """
    요청 URL의 쿼리와 서명용 쿼리를 동일한 canonical(RFC3986, 정렬)로 고정하여
    게이트웨이의 내부 재계산과 불일치를 방지합니다.
    """
    if not (COUPANG_ACCESS_KEY and COUPANG_SECRET_KEY):
        raise RuntimeError("쿠팡 API 키가 설정되지 않았습니다.")

    url_query = _cp_canonical_query_for_signature(params)
    url = f"{COUPANG_BASE_URL}{path}" + (f"?{url_query}" if url_query else "")

    headers, _signed_date, _canonical_query = _cp_signed_headers(
        method, path, params, COUPANG_ACCESS_KEY, COUPANG_SECRET_KEY, vendor_id=COUPANG_VENDOR_ID
    )

    try:
        resp = requests.request(method=method, url=url, headers=headers, timeout=30)
        resp.raise_for_status()
        return resp.json()
    except requests.HTTPError as e:
        body = ""
        try:
            body = resp.text[:1000]
        except Exception:
            pass
        msg = f"{resp.status_code} {resp.reason}\nurl={url}\nresp_body={body}"
        raise requests.HTTPError(msg, response=resp, request=resp.request) from e

# =========================
# Google Sheets 래퍼
# =========================
class SheetsClient:
    def __init__(self, json_path: str, sheet_id: str, worksheet_name: str, logger):
        self.json_path = json_path
        self.sheet_id = sheet_id
        self.worksheet_name = worksheet_name
        self.logger = logger
        self.gc = None
        self.ws = None
        self.CREATE_WORKSHEET_IF_MISSING = False

    def connect(self):
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(self.json_path, scopes=scopes)
        self.gc = gspread.authorize(creds)
        sh = self.gc.open_by_key(self.sheet_id)
        try:
            self.ws = sh.worksheet(self.worksheet_name)
            self.logger(f"✅ Google Sheets 연결 완료 (워크시트: {self.worksheet_name})")
        except gspread.WorksheetNotFound:
            titles = [w.title for w in sh.worksheets()]
            self.logger(f"⚠️ 워크시트 '{self.worksheet_name}'를 찾지 못함. 현재 탭들: {titles}")
            if self.CREATE_WORKSHEET_IF_MISSING:
                self.ws = sh.add_worksheet(title=self.worksheet_name, rows=1000, cols=30)
                self.logger(f"🆕 워크시트 생성: {self.worksheet_name}")
            else:
                raise

    def get_next_index(self) -> int:
        try:
            col_values = self.ws.col_values(1)
            last = None
            for v in reversed(col_values):
                if v.strip():
                    last = v
                    break
            if last is None:
                return 1
            return int(last) + 1 if is_int_string(last) else 1
        except Exception as e:
            self.logger(f"⚠️ A열 인덱스 계산 실패, 1로 시작: {e}")
            return 1

    def find_first_empty_row_in_col_a_from_top(self) -> int:
        values = self.ws.col_values(1)
        if not values:
            return 1
        for i, v in enumerate(values, start=1):
            if not str(v).strip():
                return i
        return len(values) + 1

    def append_row_with_retry(self, row_values, max_tries=5, base_sleep=0.6):
        attempt = 0
        while True:
            try:
                self.ws.append_row(row_values, value_input_option="USER_ENTERED")
                return True
            except gspread.exceptions.APIError as e:
                attempt += 1
                try:
                    resp = getattr(e, "response", None)
                    status = getattr(resp, "status_code", None)
                    text = getattr(resp, "text", "")
                    self.logger(f"❌ APIError(status={status}): {text[:500]}")
                except Exception:
                    self.logger(f"❌ APIError: {e}")
                if attempt >= max_tries:
                    return False
                sleep_s = base_sleep * (2 ** (attempt - 1))
                self.logger(f"⏳ 재시도 {attempt}/{max_tries} ... {sleep_s:.1f}s")
                time.sleep(sleep_s)
            except (TransportError, Exception) as e:
                attempt += 1
                self.logger(f"❌ 전송/기타 오류: {repr(e)}")
                if attempt >= max_tries:
                    return False
                sleep_s = base_sleep * (2 ** (attempt - 1))
                self.logger(f"⏳ 재시도 {attempt}/{max_tries} ... {sleep_s:.1f}s")
                time.sleep(sleep_s)

# =========================
# 메인 앱
# =========================
class ChromeCrawler(QWidget):
    clickDetected = pyqtSignal(int, int)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("크롬 크롤링 도구 (gspread 전환)")
        self.setGeometry(0, 0, 400, 550)

        # 상태값
        self.target_title = None
        self.target_window = None
        self.driver = None
        self._listener = None
        self._waiting_click = False
        self._sheet_click_wait = False
        self._click_timer = None

        # 크롤 결과
        self.crawled_title = ""
        self.crawled_price = ""
        self.crawled_url = ""

        # Google Sheets
        self.sheets = SheetsClient(SERVICE_ACCOUNT_JSON, SHEET_ID, WORKSHEET_NAME, self._log)
        self.row_index_cache = None

        # =========================
        # UI
        # =========================
        layout = QVBoxLayout()
        layout.setSpacing(6)
        layout.setContentsMargins(8, 8, 8, 8)

        self.label = QLabel("🖱 대상 윈도우: 없음")
        layout.addWidget(self.label)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        # 1) clear + Sheets 연결
        row_a = QHBoxLayout()
        self.btn_clear = QPushButton("Txt clear")
        self.btn_clear.clicked.connect(self.log.clear)
        row_a.addWidget(self.btn_clear)

        self.btn_sheets = QPushButton("Sheets 연결")
        self.btn_sheets.clicked.connect(self.connect_sheets)
        row_a.addWidget(self.btn_sheets)
        layout.addLayout(row_a)

        # 2) 크롬(디버그) + 기존 창 연결 테스트
        row_b = QHBoxLayout()
        self.btn_launch = QPushButton("크롬(디버그) 실행")
        self.btn_launch.clicked.connect(self.launch_debug_chrome)
        row_b.addWidget(self.btn_launch)

        self.btn_test = QPushButton("기존 창 연결 테스트")
        self.btn_test.clicked.connect(self.test_attach_existing)
        row_b.addWidget(self.btn_test)
        layout.addLayout(row_b)

        # 3) 대상윈도우 + 기록
        row_c = QHBoxLayout()
        self.btn_select = QPushButton("대상윈도우 (Shift+Z)")
        self.btn_select.clicked.connect(self.select_target_window)
        row_c.addWidget(self.btn_select)

        self.btn_record = QPushButton("기록 (Shift+X)")
        self.btn_record.clicked.connect(self.record_data)
        row_c.addWidget(self.btn_record)
        layout.addLayout(row_c)

        # 4) STOP + 네이버(최저가)
        row_d = QHBoxLayout()
        self.btn_stop = QPushButton("STOP (프로그램 off)")
        self.btn_stop.clicked.connect(self.close)
        row_d.addWidget(self.btn_stop)

        self.btn_health = QPushButton("네이버 (최저가))")
        self.btn_health.clicked.connect(self.naver_check)
        row_d.addWidget(self.btn_health)
        layout.addLayout(row_d)

        # 5) 쿠팡: 주문현황 + 키 확인 + API 헬스체크
        row_e = QHBoxLayout()
        self.btn_coupang = QPushButton("쿠팡 주문현황")
        self.btn_coupang.clicked.connect(self.coupang_orders)
        row_e.addWidget(self.btn_coupang)

        self.btn_cp_keycheck = QPushButton("쿠팡 키 확인")
        self.btn_cp_keycheck.clicked.connect(self.check_coupang_keys)
        row_e.addWidget(self.btn_cp_keycheck)

        self.btn_cp_health = QPushButton("쿠팡 API 헬스체크")
        self.btn_cp_health.clicked.connect(self.coupang_healthcheck)
        row_e.addWidget(self.btn_cp_health)

        layout.addLayout(row_e)

        # 버튼 높이/패딩 조정
        for btn in (
            self.btn_clear, self.btn_sheets, self.btn_launch, self.btn_test,
            self.btn_select, self.btn_record, self.btn_stop, self.btn_health,
            self.btn_coupang, self.btn_cp_keycheck, self.btn_cp_health
        ):
            btn.setMinimumHeight(28)
            btn.setStyleSheet("QPushButton { padding: 4px 8px; }")

        # 안내
        self._log(
            "ℹ️ 사용법:\n"
            "1) [Sheets 연결] → [크롬(디버그) 실행] 후 대상 페이지를 엽니다.\n"
            "2) [대상윈도우] 클릭 → 안내에 따라 '본문'을 클릭(5초 내).\n"
            "3) 크롤 완료 후 [기록]으로 시트에 반영합니다.\n"
        )

        self.setLayout(layout)

        # 단축키
        QShortcut(QKeySequence("Shift+Z"), self, activated=self.select_target_window)
        QShortcut(QKeySequence("Shift+X"), self, activated=self.record_data)

        # 전역 클릭 시그널
        self.clickDetected.connect(self._handle_click_on_main)

        # 자동 초기화
        QTimer.singleShot(300, self._startup_sequence)

    # ---------- 로깅 ----------
    def _log(self, msg: str):
        self.log.append(msg)
        print(msg)

    # ---------- 공통 HTTP 에러 로깅 ----------
    def _log_http_error(self, e: Exception, context: str = ""):
        import requests
        if isinstance(e, requests.HTTPError):
            resp = getattr(e, "response", None)
            req = getattr(e, "request", None)
            status = getattr(resp, "status_code", None)
            reason = getattr(resp, "reason", "")
            url = getattr(req, "url", "(unknown)")
            try:
                body = resp.text if resp is not None else str(e)
            except Exception:
                body = str(e)
            if context:
                self._log(f"❌ {context}: {status or 'N/A'} {reason or e.__class__.__name__}")
            else:
                self._log(f"❌ 요청 실패: {status or 'N/A'} {reason or e.__class__.__name__}")
            self._log(f"url={url}")
            self._log(f"resp_body={(body or '')[:1000]}")
        else:
            if context:
                self._log(f"❌ {context} 중 예외: {repr(e)}")
            else:
                self._log(f"❌ 예외: {repr(e)}")

    # ---------- 자동 시작 시퀀스 ----------
    def _startup_sequence(self):
        self._log("🚀 시작: 자동 초기화 시퀀스 실행")
        try:
            self.connect_sheets()
        except Exception as e:
            self._log(f"⚠️ 자동 시트 연결 실패: {e}")

        if self.sheets.ws is None:
            self._log("ℹ️ Sheets 미연결 → '기존 창 연결 테스트' 수행")
            ok = self._attach_existing_ok()
            if ok:
                self.test_attach_existing()
            else:
                self._log("ℹ️ 기존 창 연결 실패 → '크롬(디버그) 실행' 수행")
                self.launch_debug_chrome()
        else:
            self._log("✅ Sheets 연결 완료(자동)")

    def _attach_existing_ok(self) -> bool:
        try:
            if not is_port_open("127.0.0.1", DEBUGGER_PORT):
                self._log("ℹ️ 디버그 포트가 열려 있지 않음")
                return False
            driver = self._attach_driver()
            _ = driver.window_handles
            self._log("✅ 기존 창 연결 OK")
            return True
        except Exception as e:
            self._log(f"ℹ️ 기존 창 연결 실패: {e}")
            return False

    # 네이버 쇼핑 열기
    def _open_naver_shopping_with_title(self, sort_low_price: bool = True):
        try:
            title = (self.crawled_title or "").strip()
            if not title:
                self._log("ℹ️ 제목이 없어 네이버 쇼핑 검색을 생략합니다.")
                return

            driver = self._attach_driver()
            from urllib.parse import quote_plus
            base_url = "https://search.shopping.naver.com/search/all"
            q = f"query={quote_plus(title)}"
            sort = "sort=price_asc" if sort_low_price else "sort=rel"
            search_url = f"{base_url}?{q}&{sort}"

            driver.execute_script("window.open(arguments[0], '_blank');", search_url)
            driver.switch_to.window(driver.window_handles[-1])
            self._log(f"🟢 네이버 쇼핑 검색 탭 오픈(낮은가격순 시도): {search_url}")

            if not sort_low_price:
                return

            try:
                WebDriverWait(driver, 5).until(
                    lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
                )
            except Exception:
                pass

            if "sort=price_asc" in (driver.current_url or ""):
                return

            click_js = r"""
            const keywords = ['낮은가격순', '가격낮은순'];
            function clickByText(nodes) {
              for (const el of nodes) {
                try {
                  const t = (el.innerText || el.textContent || '').trim();
                  if (!t) continue;
                  for (const k of keywords) {
                    if (t.includes(k)) { el.click(); return true; }
                  }
                } catch (e) {}
              }
              return false;
            }
            const order = ['button','a','span','div','li'];
            for (const tag of order) {
              const list = document.querySelectorAll(tag);
              if (clickByText(list)) return true;
            }
            return false;
            """
            clicked = driver.execute_script(click_js)
            if clicked:
                self._log("✅ 정렬 UI 클릭으로 '낮은 가격순' 적용 시도")
                try:
                    WebDriverWait(driver, 5).until(lambda d: "price_asc" in (d.current_url or ""))
                except Exception:
                    pass
            else:
                self._log("⚠️ 정렬 UI 요소를 찾지 못했습니다. (페이지 UI 변경 가능)")
        except Exception as e:
            self._log(f"⚠️ 네이버 쇼핑 검색/정렬 처리 실패: {e}")

    # ---------- Sheets ----------
    def connect_sheets(self):
        try:
            self.sheets.connect()
        except Exception as e:
            self._log(f"❌ Sheets 연결 실패: {e}")
            raise

    def naver_check(self):
        self._open_naver_shopping_with_title(sort_low_price=True)

    # ---------- Chrome ----------
    def launch_debug_chrome(self):
        try:
            if is_port_open("127.0.0.1", DEBUGGER_PORT):
                self._log(f"ℹ️ 디버그 포트 {DEBUGGER_PORT} 이미 열림. 기존 창에 연결하세요.")
                return
            chrome_bin = None
            for p in CHROME_PATHS:
                if os.path.exists(p):
                    chrome_bin = p; break
            if chrome_bin is None:
                self._log("⚠️ Chrome 실행 파일을 찾지 못했습니다.")
                return
            Path(USER_DATA_DIR).mkdir(parents=True, exist_ok=True)
            cmd = [
                chrome_bin,
                f"--remote-debugging-port={DEBUGGER_PORT}",
                f"--user-data-dir={USER_DATA_DIR}",
                "--no-first-run", "--no-default-browser-check"
            ]
            subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
            for _ in range(25):
                if is_port_open("127.0.0.1", DEBUGGER_PORT):
                    self._log(f"✅ 디버깅 모드 Chrome 실행됨 (포트 {DEBUGGER_PORT}).")
                    return
                time.sleep(0.2)
            self._log("⚠️ 디버그 포트 연결 확인 실패")
        except Exception as e:
            self._log(f"[오류] 크롬(디버그) 실행 실패: {e}")

    def _attach_driver(self):
        if not is_port_open("127.0.0.1", DEBUGGER_PORT):
            raise RuntimeError("디버그 포트가 열려 있지 않습니다. 먼저 '크롬(디버그) 실행'을 눌러주세요.")
        if self.driver:
            return self.driver
        options = webdriver.ChromeOptions()
        options.debugger_address = f"127.0.0.1:{DEBUGGER_PORT}"
        self.driver = webdriver.Chrome(options=options)
        return self.driver

    def test_attach_existing(self):
        try:
            driver = self._attach_driver()
            tabs_info = []
            for h in driver.window_handles:
                driver.switch_to.window(h)
                tabs_info.append(f"- {safe_str(driver.title).strip()} | {safe_str(driver.current_url).strip()}")
            msg = "🔗 디버그 세션 탭 목록:\n" + ("\n".join(tabs_info) if tabs_info else "(없음)")
            self._log(msg)
        except Exception as e:
            self._log(f"[오류] 기존 창 연결 테스트 실패: {e}")

    # 시트 창 활성화
    def _bring_sheet_to_front(self):
        try:
            sheet_url_prefix = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
            if is_macos():
                osa = f'''
                tell application "Google Chrome"
                    set thePrefix to "{sheet_url_prefix}"
                    set foundWin to missing value
                    set foundIdx to -1
                    repeat with w in windows
                        set i to 0
                        repeat with t in tabs of w
                            set i to i + 1
                            if (URL of t) starts with thePrefix then
                                set foundWin to w
                                set active tab index of w to i
                                set index of w to 1
                                activate
                                return
                            end if
                        end repeat
                    end repeat
                    open location thePrefix & "/edit"
                    activate
                end tell
                '''
                subprocess.run(["osascript", "-e", osa], check=False)
            else:
                titles = []
                try:
                    titles = gw.getAllTitles()
                except Exception:
                    pass
                cand = [t for t in titles if isinstance(t, str) and ("Google Sheets" in t or "스프레드시트" in t)]
                if cand:
                    wlist = gw.getWindowsWithTitle(cand[0])
                    if wlist:
                        try:
                            wlist[0].activate()
                        except Exception:
                            pass
                try:
                    import webbrowser
                    webbrowser.open(sheet_url_prefix + "/edit", new=0, autoraise=True)
                except Exception:
                    pass
        except Exception as e:
            self._log(f"⚠️ 시트 창 활성화 실패: {e}")

    # ---------- 대상 선택 & 크롤 ----------
    def select_target_window(self):
        self._log("🖱 **크롤링할 크롬 탭의 본문**을 클릭해 주세요. (5초 내)")
        self.label.setText("🔍 본문을 클릭하세요 (주소창 X). 5초 내 미클릭 시 경고.")

        self.showMinimized()
        self._waiting_click = True
        self._sheet_click_wait = False

        if self._click_timer is None:
            self._click_timer = QTimer(self)
            self._click_timer.setSingleShot(True)
            self._click_timer.timeout.connect(self._on_click_timeout_select)
        self._click_timer.start(CLICK_TIMEOUT_MS_SELECT)

        def on_click(x, y, button, pressed):
            if pressed and self._waiting_click:
                self.clickDetected.emit(int(x), int(y))
        self._listener = MouseListener(on_click=on_click)
        self._listener.start()

    def _on_click_timeout_select(self):
        if not self._waiting_click:
            return
        self._waiting_click = False
        try:
            if self._listener: self._listener.stop()
        except Exception:
            pass
        finally:
            self._listener = None
        self._log("⏰ 5초 내 클릭이 감지되지 않았습니다. 다시 [대상윈도우]를 눌러 본문을 클릭하세요.")

    def _handle_click_on_main(self, x: int, y: int):
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

        wins_at = self._gw_get_windows_at(x, y)
        win = wins_at[0] if wins_at else None
        picked_title = safe_str(getattr(win, "title", "")) if win else ""
        if not picked_title:
            self._log("❌ 클릭 지점에서 활성 창 제목을 찾지 못했습니다. 본문 클릭/권한 확인.")
            return

        self.target_window = win
        self.target_title = picked_title
        self.label.setText(f"🎯 대상 윈도우: {self.target_title}")

        self.showNormal(); self.raise_(); self.activateWindow()
        self.crawl_data()

    def _gw_get_windows_at(self, x: int, y: int):
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

    def crawl_data(self):
        if not self.target_title:
            self._log("⚠️ 대상 탭이 선택되지 않았습니다.")
            return
        try:
            if self.target_window:
                try:
                    self.target_window.activate(); time.sleep(0.2)
                except Exception:
                    pass

            driver = self._attach_driver()

            self._log("🧭 탭 매칭: URL패턴 → 제목 포함")
            end_time = time.time() + 5.0
            target_handle = None

            if URL_PATTERNS:
                while time.time() < end_time and not target_handle:
                    for h in driver.window_handles:
                        driver.switch_to.window(h)
                        if any(p in (driver.current_url or "") for p in URL_PATTERNS):
                            target_handle = h; break
                    if not target_handle:
                        time.sleep(0.2)

            if not target_handle:
                end_time2 = time.time() + 5.0
                want = safe_str(self.target_title).strip()
                while time.time() < end_time2 and not target_handle:
                    for h in driver.window_handles:
                        driver.switch_to.window(h)
                        if want and want in safe_str(driver.title).strip():
                            target_handle = h; break
                    if not target_handle:
                        time.sleep(0.2)

            if not target_handle:
                self._log("❌ 5초 내 '대상 탭'을 찾지 못했습니다.")
                return

            driver.switch_to.window(target_handle)

            current_url = safe_str(driver.current_url).strip()
            self.crawled_url = current_url
            self._log(f"🔗 URL: {current_url}")

            blocked = ("chrome://", "chrome-extension://", "edge://", "about:", "data:")
            if any(current_url.startswith(s) for s in blocked) or current_url.lower().endswith(".pdf"):
                self._log("❌ 이 페이지는 DOM 접근이 제한됩니다.")
                return

            try:
                WebDriverWait(driver, 3).until(
                    lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
                )
            except Exception:
                pass

            title_value = ""
            wait = WebDriverWait(driver, 5)
            for sel in selectors_for_url(current_url):
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
            self._log(f"🟢 제목: {self.crawled_title or '(없음)'}")

            price_digits = ""
            wait_p = WebDriverWait(driver, 3)
            for sel in price_selectors_for_url(current_url):
                try:
                    el = wait_p.until(EC.visibility_of_element_located((By.CSS_SELECTOR, sel)))
                    txt = (el.text or "").strip()
                    if not txt:
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
            self._log(f"💰 가격(숫자만): {self.crawled_price or '(없음)'}")

            self._log("—" * 40)
            self._log(f"제목: {self.crawled_title or '(없음)'}")
            self._log(f"가격(숫자만): {self.crawled_price or '(없음)'}")
            self._log(f"URL: {self.crawled_url or '(없음)'}")
            self._log("—" * 40)

            self._log("📝 크롤 완료: 시트에 바로 기록합니다.")
            self.record_data()

        except Exception as e:
            self._log(f"[오류] 크롤링 실패: {e}")

    # ---------- 시트 기록(핵심) ----------
    def _write_row_to_first_empty_a(self):
        if self.sheets.ws is None:
            self._log("⚠️ 먼저 [Sheets 연결]을 눌러 구글시트에 연결해 주세요.")
            return

        target_row = self.sheets.find_first_empty_row_in_col_a_from_top()

        COLS = {c:i for i,c in enumerate(
            ["A","B","C","D","E","F","G","H","I","J",
             "K","L","M","N","O","P","Q","R","S","T",
             "U","V","W","X","Y"], start=1)}

        a_index = str(self.sheets.get_next_index())

        row_buffer = [""] * 25
        row_buffer[COLS["A"]-1] = a_index
        row_buffer[COLS["B"]-1] = today_fmt()
        row_buffer[COLS["C"]-1] = label_for_domain(self.crawled_url)
        row_buffer[COLS["F"]-1] = self.crawled_title or ""
        row_buffer[COLS["H"]-1] = self.crawled_price or ""
        row_buffer[COLS["I"]-1] = FIXED_CONST_FEE
        row_buffer[COLS["J"]-1] = f"=H{target_row}+I{target_row}"
        row_buffer[COLS["K"]-1] = "10.8%"
        row_buffer[COLS["M"]-1] = f"=J{target_row}+(R{target_row}*(K{target_row}*1.1))"
        row_buffer[COLS["N"]-1] = f"=O{target_row}/R{target_row}"
        row_buffer[COLS["O"]-1] = f"=R{target_row}-M{target_row}+K{target_row}-P{target_row}+L{target_row}"
        row_buffer[COLS["R"]-1] = f"=Q{target_row}"
        row_buffer[COLS["S"]-1] = f"=R{target_row}/1.1"
        row_buffer[COLS["T"]-1] = f"=S{target_row}*1.1-S{target_row}"
        row_buffer[COLS["V"]-1] = self.crawled_url or ""

        rng = f"A{target_row}:Y{target_row}"
        self.sheets.ws.update(values=[row_buffer], range_name=rng, value_input_option="USER_ENTERED")
        self._log(f"✅ 행 {target_row} (A..Y)에 기록 완료")

        # 👉 URL을 클립보드로 복사
        try:
            if self.crawled_url:
                pyperclip.copy(self.crawled_url)
                self._log("📋 현재 상품 URL을 클립보드에 복사했습니다.")
        except Exception as e:
            self._log(f"⚠️ 클립보드 복사 실패: {e}")

        self._bring_sheet_to_front()

    def record_data(self):
        if not self.crawled_url:
            self._log("⚠️ 먼저 [대상윈도우]로 제목/가격/URL을 크롤링해 주세요.")
            return
        try:
            self._write_row_to_first_empty_a()
        except Exception as e:
            self._log(f"[오류] 시트 기록 실패: {e}")

    # ---------- 시트 클릭 대기 → 기록 ----------
    def _wait_for_sheet_click_then_write(self):
        if self._sheet_click_wait:
            return

        self._sheet_click_wait = True
        start_ts = time.time()
        self._log("⌛ 시트 클릭 대기 시작 (10초)")

        def wait_click():
            nonlocal start_ts
            with mouse.Events() as events:
                for event in events:
                    if (time.time() - start_ts) * 1000 >= CLICK_TIMEOUT_MS_RECORD:
                        break
                    if isinstance(event, mouse.Events.Click) and event.pressed:
                        self._sheet_click_wait = False
                        self.record_data()
                        return
            self._sheet_click_wait = False
            self._log("⏰ 10초 내 시트 클릭이 감지되지 않았습니다. [기록] 버튼으로 입력하세요.")

        import threading
        t = threading.Thread(target=wait_click, daemon=True)
        t.start()

    # ==== 쿠팡 주문조회 + 시트기록 ====
    def _fetch_coupang_orders(self) -> list[dict]:
        if not (COUPANG_VENDOR_ID and COUPANG_ACCESS_KEY and COUPANG_SECRET_KEY):
            self._log("❌ 쿠팡 API 키/벤더ID 설정이 비어 있습니다.")
            return []

        from datetime import datetime, timedelta, timezone
        to_dt = datetime.now(timezone.utc)
        from_dt = to_dt - timedelta(days=COUPANG_LOOKBACK_DAYS)
        created_from = from_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        created_to = to_dt.strftime("%Y-%m-%dT%H:%M:%SZ")

        path = f"/v2/providers/openapi/apis/api/v4/vendors/{COUPANG_VENDOR_ID}/ordersheets"
        all_rows: list[dict] = []

        for st in CP_QUERY_STATUSES:
            next_token = None
            while True:
                params = {
                    "createdAtFrom": created_from,
                    "createdAtTo": created_to,
                    "status": st,
                    "maxPerPage": 50,
                }
                if next_token:
                    params["nextToken"] = next_token

                try:
                    data = _cp_request("GET", path, params)
                except requests.HTTPError as e:
                    self._log_http_error(e, context=f"쿠팡 API 호출 실패(status={st})")
                    break
                except Exception as e:
                    self._log(f"⚠️ 쿠팡 API 호출 실패(status={st}): {repr(e)}")
                    break

                result_code = str(data.get("code", "")).upper()
                if result_code not in ("SUCCESS", "OK", "200"):
                    self._log(f"⚠️ 응답 코드 이상(status={st}): {data.get('message')}")
                    break

                datas = data.get("data") or {}
                sheets = datas.get("orderSheets") or datas.get("shipmentBoxInfos") or []
                for sheet in sheets:
                    order_id = sheet.get("orderId") or sheet.get("orderIdMask") or ""
                    order_date = sheet.get("orderedAt") or sheet.get("orderDate") or ""
                    receiver = (sheet.get("receiver") or {})
                    recv_name = receiver.get("name", "")
                    recv_addr = receiver.get("addr1", "")
                    recv_phone = receiver.get("contact1", "") or receiver.get("contact2", "")

                    items = sheet.get("orderItems") or []
                    for it in items:
                        item_name = it.get("vendorItemName") or it.get("sellerProductName") or it.get("productName") or ""
                        order_item_id = it.get("orderItemId") or it.get("vendorItemId") or ""
                        qty = it.get("quantity") or it.get("shippingCount") or 1
                        paid_price = it.get("paidPrice") or it.get("unitPrice") or 0
                        tracking_no = it.get("invoiceNumber") or it.get("trackingNumber") or ""
                        carrier = it.get("deliveryCompanyName") or it.get("deliveryCompanyCode") or ""
                        status_text = CP_STATUS_MAP.get(st, st)

                        all_rows.append({
                            "주문일시": order_date,
                            "상태": status_text,
                            "주문번호": order_id,
                            "주문아이템ID": order_item_id,
                            "상품명": item_name,
                            "수량": qty,
                            "결제금액": paid_price,
                            "수취인": recv_name,
                            "연락처": recv_phone,
                            "주소": recv_addr,
                            "송장번호": tracking_no,
                            "택배사": carrier,
                        })

                next_token = datas.get("nextToken")
                if not next_token:
                    break

        self._log(f"📦 쿠팡 주문 수집 완료: {len(all_rows)}건")
        return all_rows

    def _write_coupang_orders_to_sheet(self, rows: list[dict]):
        if self.sheets.ws is None:
            self._log("⚠️ Sheets 연결이 필요합니다. 먼저 [Sheets 연결] 버튼을 눌러주세요.")
            return

        # 워크시트 열기/없으면 생성
        try:
            ws = self.sheets.gc.open_by_key(SHEET_ID).worksheet(COUPANG_WS_NAME)
        except gspread.WorksheetNotFound:
            ws = self.sheets.gc.open_by_key(SHEET_ID).add_worksheet(title=COUPANG_WS_NAME, rows=2000, cols=30)

        if not rows:
            headers = ["주문일시","상태","주문번호","주문아이템ID","상품명","수량","결제금액","수취인","연락처","주소","송장번호","택배사"]
            ws.clear()
            ws.update(values=[headers], range_name="A1:L1")
            self._log("ℹ️ 쿠팡 주문 데이터가 없어 헤더만 갱신했습니다.")
            return

        headers = list(rows[0].keys())
        values = [headers] + [[str(r.get(h, "")) for h in headers] for r in rows]

        ws.clear()
        end_col_index = len(headers)
        end_col_letter = chr(ord('A') + end_col_index - 1)
        end_row = len(values)
        rng = f"A1:{end_col_letter}{end_row}"
        ws.update(values=values, range_name=rng, value_input_option="USER_ENTERED")
        self._log(f"✅ '{COUPANG_WS_NAME}' 탭에 {len(rows)}건 업데이트 완료")

    # === 쿠팡 주문현황 버튼 동작 ===
    def coupang_orders(self):
        if self.sheets.ws is None:
            self._log("ℹ️ Sheets 미연결: 자동으로 연결 시도합니다.")
            try:
                self.connect_sheets()
            except Exception as e:
                self._log(f"❌ Sheets 연결 실패: {e}")
                return

        try:
            rows = self._fetch_coupang_orders()
        except Exception as e:
            self._log(f"❌ 쿠팡 주문 조회 중 오류: {e}")
            return

        try:
            self._write_coupang_orders_to_sheet(rows)
        except Exception as e:
            self._log(f"❌ 쿠팡 주문 기록 중 오류: {e}")

    # === 쿠팡 키 확인 버튼 동작 ===
    def check_coupang_keys(self):
        """쿠팡 키 JSON 파일을 로드해서 값 유효성/마스킹 출력 + 간단 HMAC 서명 생성 테스트"""
        try:
            p = Path(COUPANG_KEYS_JSON)
            if not p.exists():
                self._log(f"❌ 키 파일을 찾지 못했습니다: {COUPANG_KEYS_JSON}")
                self._log("➡ 경로/파일명을 다시 확인하거나 JSON을 생성해 주세요.")
                return

            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)

            vendor_id = (data.get("vendor_id") or "").strip()
            access_key = (data.get("access_key") or "").strip()
            secret_key = (data.get("secret_key") or "").strip()

            self._log("✅ JSON 파일 읽기 성공")
            self._log(f"• Vendor ID: {vendor_id or '(빈 값)'}")
            self._log(f"• Access Key: {access_key or '(빈 값)'}")
            self._log(f"• Secret Key: {_mask(secret_key) if secret_key else '(빈 값)'}")

            problems = []
            if not vendor_id: problems.append("vendor_id가 비어 있습니다.")
            if not access_key: problems.append("access_key가 비어 있습니다.")
            if not secret_key: problems.append("secret_key가 비어 있습니다.")
            if problems:
                for m in problems:
                    self._log(f"⚠️ {m}")
                return

            mismatches = []
            if COUPANG_VENDOR_ID != vendor_id:
                mismatches.append("전역 Vendor ID와 JSON의 vendor_id가 다릅니다.")
            if COUPANG_ACCESS_KEY != access_key:
                mismatches.append("전역 Access Key와 JSON의 access_key가 다릅니다.")
            if COUPANG_SECRET_KEY != secret_key:
                mismatches.append("전역 Secret Key와 JSON의 secret_key가 다릅니다.")
            if mismatches:
                self._log("⚠️ 전역 설정과 JSON 파일의 값이 일치하지 않습니다:")
                for m in mismatches:
                    self._log(f"   - {m}")
                self._log("➡ JSON을 수정했으면 프로그램을 재시작하거나, 상단 상수 경로/로딩 부분을 확인하세요.")

            # --- HMAC 서명 생성 테스트 (API 호출은 하지 않음) ---
            try:
                test_path = f"/v2/providers/openapi/apis/api/v4/vendors/{vendor_id}/ordersheets"
                test_query = "status=ACCEPT&maxPerPage=50"
                headers, _sd, _cq = _cp_signed_headers("GET", test_path, test_query, access_key, secret_key, vendor_id=vendor_id)
                self._log("🔐 HMAC 서명 생성 테스트 성공")
                auth = headers.get("Authorization", "")
                if auth:
                    self._log(f"• Authorization 헤더 앞부분: {auth[:60]}...")
            except Exception as e:
                self._log(f"❌ HMAC 서명 생성 실패: {e}")

            self._log("🟢 키 확인 완료")

        except json.JSONDecodeError as e:
            self._log(f"❌ JSON 파싱 실패: {e}")
            self._log("➡ 파일 내용이 유효한 JSON 형식인지 확인하세요.")
        except Exception as e:
            self._log(f"❌ 키 확인 중 오류: {e}")

    # === 쿠팡 API 헬스체크 버튼 동작 ===
    def coupang_healthcheck(self):
        """
        쿠팡 OpenAPI 헬스체크:
        1) ordersheets(벤더스코프, v4)로 1차 시도
        2) 판매자상품 목록 경로들로 폴백 시도
        """
        import requests
        from datetime import datetime, timedelta, timezone

        self._log("🩺 쿠팡 API 헬스체크 시작")

        if not (COUPANG_VENDOR_ID and COUPANG_ACCESS_KEY and COUPANG_SECRET_KEY):
            self._log("❌ 쿠팡 키/벤더ID가 비어 있습니다. coupang_keys.json 확인")
            return

        # 1) ordersheets 우선 시도 (최근 1일, ACCEPT, 1건만)
        try:
            to_dt = datetime.now(timezone.utc)
            from_dt = to_dt - timedelta(days=1)
            params = {
                "createdAtFrom": from_dt.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "createdAtTo": to_dt.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "status": "ACCEPT",
                "maxPerPage": 1,
            }
            path = f"/v2/providers/openapi/apis/api/v4/vendors/{COUPANG_VENDOR_ID}/ordersheets"
            data = _cp_request("GET", path, params)
            code = str(data.get("code", "")).upper()
            self._log(f"✅ 헬스체크 성공: path='{path}', params={params} (code={code or 'N/A'})")
            self._log("🟢 쿠팡 API 키/서명/경로 정상으로 보입니다.")
            return
        except requests.HTTPError as e:
            self._log_http_error(e, context="헬스체크(ordersheets) 실패")
        except Exception as e:
            self._log(f"❌ 헬스체크(ordersheets) 중 예외: {repr(e)}")

        # 2) 판매자상품 목록 경로 폴백
        path_candidates = [
            f"/v2/providers/openapi/apis/api/v1/vendors/{COUPANG_VENDOR_ID}/seller-products",
            f"/v2/providers/openapi/apis/api/v2/vendors/{COUPANG_VENDOR_ID}/seller-products",
            "/v2/providers/openapi/apis/api/v1/marketplace/seller-products",
        ]
        param_candidates = [
            {"page": 0, "size": 1},
            {"maxPerPage": 1},
        ]

        last_err = None
        for path in path_candidates:
            for params in param_candidates:
                try:
                    data = _cp_request("GET", path, params)
                    code = str(data.get("code", "")).upper()
                    self._log(f"✅ 헬스체크 성공: path='{path}', params={params} (code={code or 'N/A'})")

                    payload = data.get("data")
                    if isinstance(payload, dict):
                        items = payload.get("content") or payload.get("sellerProductList") or payload.get("items") or []
                    else:
                        items = payload or []

                    cnt = len(items) if isinstance(items, list) else (1 if items else 0)
                    self._log(f"📦 샘플 항목 수: {cnt}건")
                    if cnt:
                        first = items[0]
                        name = (isinstance(first, dict) and (
                            first.get("sellerProductName") or first.get("productName") or first.get("vendorItemName")
                        )) or "(이름 없음)"
                        any_id = (isinstance(first, dict) and (
                            first.get("sellerProductId") or first.get("productId") or first.get("vendorItemId")
                        )) or None
                        self._log(f"🔎 첫 항목: name='{name}', id={any_id}")

                    self._log("🟢 쿠팡 API 키/서명/경로 정상으로 보입니다.")
                    return

                except requests.HTTPError as e:
                    self._log_http_error(e, context=f"헬스체크 실패(path='{path}', params={params})")
                    resp = getattr(e, "response", None)
                    status = getattr(resp, "status_code", None)
                    body = ""
                    try:
                        body = resp.text or ""
                    except Exception:
                        pass

                    if status == 404 or "No exactly matching API specification" in body:
                        last_err = e
                        continue
                    if status in (401, 403):
                        self._log(
                            "🔎 점검 항목:\n"
                            "  1) 키가 '판매자센터(Wing) OpenAPI' 키인지 (파트너스 키 아님)\n"
                            "  2) 판매자센터 > 시스템 연동 > Open API: 사용 '활성' 및 권한 승인\n"
                            "  3) (필요 시) 허용 IP 목록에 현재 PC 공인 IP 등록\n"
                            "  4) PC 시간 자동 동기화(UTC, 수 초 이하 오차)\n"
                        )
                        return
                    last_err = e
                    continue
                except Exception as e:
                    self._log(f"❌ 헬스체크 중 예외(path='{path}', params={params}): {repr(e)}")
                    last_err = e
                    continue

        self._log("❌ 모든 경로/파라미터 후보가 실패했습니다. 아래 마지막 오류를 참고하세요.")
        if last_err:
            if isinstance(last_err, requests.HTTPError):
                self._log_http_error(last_err, context="헬스체크 최종 실패")
            else:
                self._log(f"최종 예외: {repr(last_err)}")

# =========================
# 엔트리 포인트
# =========================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ChromeCrawler()
    win.show()
    sys.exit(app.exec())
