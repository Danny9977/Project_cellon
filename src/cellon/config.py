# config.py
from pathlib import Path
import platform
from datetime import datetime

# ====== category_ai 에서 카테고리 엑셀 파일 경로 ==========
CATEGORY_EXCEL_DIR = Path.home() / "Desktop" / "category_excels"

# ====== category_ai 캐시 디렉토리 (카테고리 마스터/파일 캐시) ======
# 원하시는 위치로 바꿔도 됩니다. (예: 프로젝트 내부 cache 폴더 등)
CACHE_DIR = Path.home() / "Desktop" / "category_cache"
# 예를 들어 프로젝트 폴더 안에 두고 싶으면:
# PROJECT_ROOT = Path(__file__).resolve().parent
# CACHE_DIR = PROJECT_ROOT / "cache"

# === 로컬 LLM(Ollama) 설정 ===
LOCAL_LLM_BASE_URL = "http://localhost:11434"
LOCAL_LLM_MODEL = "llama3:8b"
#LOCAL_LLM_MODEL = "phi3:medium"


# =========================
# Google Sheets 설정
# =========================
SERVICE_ACCOUNT_JSON = "/Users/jeehoonkim/Desktop/api/google_api/service_account.json"
SHEET_ID = "1OEg01RdJyesSy7iQSEyQHdYpCX5MSsNUfD0lkUYq8CM"
WORKSHEET_NAME = "소싱상품목록"

# 쿠팡 주문현황 시트 이름
COUPANG_WS_NAME = "쿠팡주문현황"

# =========================
# 크롬 디버그/브라우저 설정
# =========================
DEBUGGER_ADDR = "127.0.0.1:9222"
DEBUGGER_PORT = 9222

CHROME_PATHS = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Google Chrome Beta.app/Contents/MacOS/Google Chrome Beta",
    "/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary",
]

USER_DATA_DIR = str(Path("/Users/Shared/chrome_dev"))

# =========================
# UI/지연/타임아웃 설정
# =========================
CLICK_TIMEOUT_MS_SELECT = 5000   # 대상윈도우 선택(본문 클릭) 대기 타임아웃
CLICK_TIMEOUT_MS_RECORD = 10000  # 시트 클릭 대기 타임아웃
KEY_DELAY_SEC = 0.01
CLICK_STABILIZE_SEC = 0.01
NAV_DELAY_SEC = 0.005

# =========================
# 날짜/수수료 기본값
# =========================
DATE_FORMAT = "M/D"        # 날짜 포맷
FIXED_CONST_FEE = "3000"   # I열 고정 수수료

# =========================
# 도메인/라벨 매핑
# =========================
DOMAIN_LABELS = {
    "domeme.domeggook.com": "도매매",
    "naver.com": "네이버",
    "costco.co.kr": "코코",
    "ownerclan.com": "오너",
}

# =========================
# 사이트별 셀렉터
# =========================
SITE_SELECTORS = {
    "domeme.domeggook.com": ["#lInfoItemTitle", "h1#lInfoItemTitle", "h1"],
    "costco.co.kr": [".product-detail__name", "h1.product-detail__name", "h1"],
}

SITE_PRICE_SELECTORS = {
    "domeme.domeggook.com": ["#lItemPrice", ".lItemPrice", "#lItemPriceText"],
}

DEFAULT_SELECTORS = [
    "#lInfoItemTitle", "h1.l.infoItemTitle",
    "h1#l\\.infoItemTitle", "h1",
    "[role=\"heading\"][aria-level=\"1\"]",
]

# =========================
# URL 패턴 (탭 매칭용)
# =========================
URL_PATTERNS = [
    "domeme.domeggook.com/s/",
    "domeme.domeggook.com",
    "costco.co.kr",
]

# =========================
# 쿠팡 Open API 설정
# =========================
COUPANG_BASE_URL = "https://api-gateway.coupang.com"
COUPANG_KEYS_JSON = str(Path("/Users/jeehoonkim/Desktop/api/coupang_api/coupang_keys.json"))

# 조회 기본 기간(일)
DEFAULT_LOOKBACK_DAYS = 7

# 조회/표시할 상태: 결제완료 → 상품준비중 → 배송지시 → 배송중 → 배송완료
CP_QUERY_STATUSES = ["ACCEPT", "INSTRUCT", "DEPARTURE", "DELIVERING", "DELIVERED"]

# 시트에 적을 한글 상태 라벨
CP_STATUS_MAP = {
    "ACCEPT":     "결제완료",
    "INSTRUCT":   "상품준비중",
    "DEPARTURE":  "배송지시",
    "DELIVERING": "배송중",
    "DELIVERED":  "배송완료",
}

# API별 상태 이름 별칭
ORDER_STATUS_ALIASES = {
    "ACCEPT":     ["ACCEPT", "PAID", "PAYMENT_COMPLETED", "ORDER_COMPLETE"],
    "INSTRUCT":   ["INSTRUCT", "READY", "READY_FOR_DELIVERY", "PREPARE_SHIPMENT"],
    "DEPARTURE":  ["DEPARTURE", "DELIVERY_REQUESTED", "SHIPPING_READY"],
    "DELIVERING": ["DELIVERING"],
    "DELIVERED":  ["DELIVERED", "DELIVERY_COMPLETED", "DONE", "FINAL_DELIVERY"],
}

STATUS_ORDER = {
    "결제완료": 0,
    "상품준비중": 1,
    "배송지시": 2,
    "배송중":   3,
    "배송완료": 4,
}

# =========================
# 코스트코 → 쿠팡 대량등록 엑셀
# =========================
SELLERTOOL_XLSM_PATH = "/Users/jeehoonkim/Desktop/Python_Project/Cellon_Project/crawling_temp/sellertool_upload.xlsm"
SELLERTOOL_SHEET_NAME = "data"  # 실제 시트 이름



def is_macos() -> bool:
    """맥OS인지 여부 확인"""
    return platform.system() == "Darwin"

def today_fmt(fmt: str = "%Y-%m-%d") -> str:
    """오늘 날짜를 지정 포맷 문자열로 반환"""
    return datetime.now().strftime(fmt)

def digits_only(s) -> str:
    """문자열에서 숫자만 뽑아냄 ('1,234원' -> '1234')"""
    if s is None:
        return ""
    return "".join(ch for ch in str(s) if ch.isdigit())

def is_int_string(s) -> bool:
    """정수로 해석 가능한 문자열인지 체크"""
    try:
        int(str(s))
        return True
    except (TypeError, ValueError):
        return False

def label_for_domain(domain: str) -> str:
    """
    도메인 문자열(예: '코코', '도매매', '네이버', '오너' 등)에 따라
    시트에 쓸 라벨 텍스트를 정함.
    """
    domain = (domain or "").strip()
    # 예전 단일파일에서 쓰던 매핑 그대로 옮기면 됨
    if "코코" in domain or "코스트코" in domain:
        return "코코"
    if "도매매" in domain:
        return "도매매"
    if "네이버" in domain:
        return "네이버"
    if "오너" in domain:
        return "오너"
    return "기타"

def _a1_col(col_index: int) -> str:
    """
    1 -> 'A', 2 -> 'B', ... 27 -> 'AA' 같은 A1 표기용 컬럼 문자열로 변환.
    예전 코드에도 동일한 로직 있었음.
    """
    result = ""
    n = col_index
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(ord('A') + rem) + result
    return result

