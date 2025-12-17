# config.py
from pathlib import Path
import platform
from datetime import datetime

# =========================
# 프로젝트 루트/공통 경로
# =========================
# config.py 파일 위치: .../Cellon_Project/src/cellon/config.py
CELLON_PACKAGE_DIR = Path(__file__).resolve().parent          # .../src/cellon
SRC_DIR = CELLON_PACKAGE_DIR.parent                           # .../src
PROJECT_ROOT = SRC_DIR.parent                                 # .../Cellon_Project
ASSETS_DIR = PROJECT_ROOT / "assets"

# ====== category_ai 에서 카테고리 엑셀 파일 경로 ==========
# 예: Cellon_Project/assets/category_excels/주방용품.xlsx ...
CATEGORY_EXCEL_DIR = ASSETS_DIR / "category_excels"

# ====== category_ai 캐시 디렉토리 (카테고리 마스터/파일 캐시) ======
# 예: Cellon_Project/assets/cache/category_master.pkl
CACHE_DIR = ASSETS_DIR / "cache"

# ====== 크롤링 / 이미지 / 업로드 관련 경로 ======
# 예: Cellon_Project/assets/temp/crawling_temp/ ...
CRAWLING_TEMP_DIR = ASSETS_DIR / "crawling_temp"
CRAWLING_TEMP_IMAGE_DIR = CRAWLING_TEMP_DIR / "image"   # 캡처 이미지
CRAWLING_TEMP_EXCEL_DIR = CRAWLING_TEMP_DIR / "excel"   # 엑셀 (있다면)

# ▶ 추가: 쿠팡 셀러툴 업로드 폼 루트
COUPANG_UPLOAD_FORM_DIR = CRAWLING_TEMP_DIR / "coupang_upload_form"

# 업로드용 폴더
UPLOAD_READY_DIR = CRAWLING_TEMP_DIR / "upload_ready"


# =========== test ============
# ✅ 추가: 쿠팡 업로드 템플릿 인덱스 JSON 경로
COUPANG_UPLOAD_INDEX_JSON = CACHE_DIR / "coupang_upload_index.json"
SELLERTOOL_SHEET_NAME = "data"

# 배경 이미지 (지금 쓰는 1000x1000)
PRODUCT_BG_IMAGE_PATH = ASSETS_DIR / "image" / "bg" / "product_bg_1000.jpg"

# 코스트코→쿠팡 대량등록용 템플릿 엑셀
# 실제 파일명이 다르면 이 한 줄만 파일명에 맞게 바꿔주세요.
# 추후 쿠팡 업로드용 매크로 엑셀 (미리 정의만)
SELLERTOOL_XLSM_PATH = UPLOAD_READY_DIR / "sellertool_upload.xlsm"
SELLERTOOL_SHEET_NAME = "data"  # 실제 시트 이름


# === 로컬 LLM(Ollama) 설정 ===
LOCAL_LLM_BASE_URL = "http://localhost:11434"
LOCAL_LLM_MODEL = "llama3:8b"
# LOCAL_LLM_MODEL = "phi3:medium"


# =========================
# Google Sheets 설정
# =========================
# 예: Cellon_Project/assets/api/google_api/service_account.json
SERVICE_ACCOUNT_JSON = ASSETS_DIR / "api" / "google_api" / "service_account.json"

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

# 크롬 프로필 디렉토리 (지금처럼 Shared 경로 사용)
USER_DATA_DIR = str(Path("/Users/Shared/chrome_dev"))


# =========================
# UI/지연/타임아웃 설정
# =========================
# 대상윈도우 선택(본문 클릭) 대기 타임아웃
CLICK_TIMEOUT_MS_SELECT = 5000
# 시트 클릭 대기 타임아웃
CLICK_TIMEOUT_MS_RECORD = 10000

KEY_DELAY_SEC = 0.01
CLICK_STABILIZE_SEC = 0.01
NAV_DELAY_SEC = 0.005


# =========================
# 날짜/수수료 기본값
# =========================
DATE_FORMAT = "M/D"   # 날짜 포맷
FIXED_CONST_FEE = "3000"  # I열 고정 수수료


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
    "#lInfoItemTitle",
    "h1.l.infoItemTitle",
    "h1#l\\.infoItemTitle",
    "h1",
    '[role="heading"][aria-level="1"]',
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

# 예: Cellon_Project/assets/api/coupang_api/coupang_keys.json
COUPANG_KEYS_JSON = ASSETS_DIR / "api" / "coupang_api" / "coupang_keys.json"

# 조회 기본 기간(일)
DEFAULT_LOOKBACK_DAYS = 7


# =========================
# 유틸 함수들
# =========================
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
    1 -> 'A', 2 -> 'B', ... 27 -> 'AA' 같은
    A1 표기용 컬럼 문자열로 변환.
    """
    result = ""
    n = col_index
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(ord("A") + rem) + result
    return result
