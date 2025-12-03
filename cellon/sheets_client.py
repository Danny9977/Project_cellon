import json
import hmac
import hashlib
import requests
import time
from datetime import datetime, timezone
from urllib.parse import urlencode
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from .config import digits_only, is_int_string

# config에서 필요한 값 import
from .config import (
    COUPANG_KEYS_JSON, COUPANG_BASE_URL, SERVICE_ACCOUNT_JSON, SHEET_ID, WORKSHEET_NAME,
    CP_QUERY_STATUSES, ORDER_STATUS_ALIASES, CP_STATUS_MAP, STATUS_ORDER, COUPANG_WS_NAME,
    FIXED_CONST_FEE, DEFAULT_LOOKBACK_DAYS
)

def extract_money_amount(m: dict | None) -> int:
    if not isinstance(m, dict):
        return 0
    units = m.get("units", 0)
    nanos = m.get("nanos", 0)
    try:
        units = int(units)
    except Exception:
        units = 0
    try:
        nanos = int(nanos)
    except Exception:
        nanos = 0
    if nanos:
        return units + round(nanos / 1_000_000_000)
    return units

def extract_paid_price_from_item(it: dict) -> int:
    if not isinstance(it, dict):
        return 0
    op = it.get("orderPrice")
    if isinstance(op, dict):
        v = extract_money_amount(op)
        if v:
            return v
    if op is not None and not isinstance(op, dict):
        s = digits_only(op)
        if s:
            try:
                return int(s)
            except Exception:
                pass
    sales = it.get("salesPrice")
    sales_val = 0
    if isinstance(sales, dict):
        sales_val = extract_money_amount(sales)
    elif sales is not None:
        s = digits_only(sales)
        if s:
            try:
                sales_val = int(s)
            except Exception:
                sales_val = 0
    qty = it.get("shippingCount") or it.get("quantity") or 1
    try:
        qty = int(qty)
    except Exception:
        qty = 1
    if sales_val and qty:
        return sales_val * qty
    for key in ("paidPrice", "paymentAmount", "price"):
        if key in it and it[key] is not None:
            s = digits_only(it[key])
            if s:
                try:
                    return int(s)
                except Exception:
                    pass
    return 0

def _cp_build_query(params: dict | None) -> str:
    if not params:
        return ""
    return urlencode(params, doseq=True)

def _cp_signed_headers_v2(method: str, path: str, sign_query: str,
                          access_key: str, secret_key: str,
                          *, signed_date: str | None = None, vendor_id: str | None = None) -> dict:
    if signed_date is None:
        signed_date = datetime.now(timezone.utc).strftime("%y%m%dT%H%M%SZ")
    message = f"{signed_date}{method.upper()}{path}{sign_query}"
    signature = hmac.new(
        secret_key.encode("utf-8"),
        message.encode("utf-8"),
        hashlib.sha256
    ).hexdigest()
    authorization = (
        f"CEA algorithm=HmacSHA256, access-key={access_key}, "
        f"signed-date={signed_date}, signature={signature}"
    )
    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "Authorization": authorization,
    }
    if vendor_id:
        headers["X-Requested-By"] = vendor_id
    return headers

def _cp_request(method: str, path: str, params: dict | None) -> dict:
    try:
        with open(COUPANG_KEYS_JSON, "r", encoding="utf-8") as f:
            coupang_keys = json.load(f)
            COUPANG_VENDOR_ID = (coupang_keys.get("vendor_id") or "").strip()
            COUPANG_ACCESS_KEY = (coupang_keys.get("access_key") or "").strip()
            COUPANG_SECRET_KEY = (coupang_keys.get("secret_key") or "").strip()
    except Exception as e:
        raise RuntimeError(f"쿠팡 키 파일을 불러오지 못했습니다: {e}")
    if not (COUPANG_ACCESS_KEY and COUPANG_SECRET_KEY):
        raise RuntimeError("쿠팡 API 키가 설정되지 않았습니다.")
    url_query = _cp_build_query(params)
    url = f"{COUPANG_BASE_URL}{path}" + (f"?{url_query}" if url_query else "")
    try:
        headers = _cp_signed_headers_v2(
            method, path, url_query, COUPANG_ACCESS_KEY, COUPANG_SECRET_KEY,
            vendor_id=COUPANG_VENDOR_ID
        )
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

def _build_ordersheets_params(date_from_utc: datetime, date_to_utc: datetime, status: str, max_per_page: int = 50):
    d_from = date_from_utc.strftime("%Y-%m-%d")
    d_to   = date_to_utc.strftime("%Y-%m-%d")
    primary = {
        "createdAtFrom": d_from,
        "createdAtTo": d_to,
        "status": status,
        "maxPerPage": max_per_page,
    }
    fallback = {
        "startTime": d_from,
        "endTime": d_to,
        "status": status,
        "maxPerPage": max_per_page,
    }
    return [primary, fallback]

def _try_ordersheets_with_variants(path: str, param_variants: list[dict]) -> dict:
    last_err = None
    for params in param_variants:
        try:
            return _cp_request("GET", path, params)
        except requests.HTTPError as e:
            resp = getattr(e, "response", None)
            status = getattr(resp, "status_code", None)
            body = ""
            try:
                body = (resp.text or "")[:500]
            except Exception:
                pass
            if status == 400 and "yyyy-MM-dd" in body:
                last_err = e
                continue
            raise
        except Exception as e:
            last_err = e
            continue
    if last_err:
        raise last_err
    raise RuntimeError("ordersheets 호출 시도 실패: 유효한 파라미터 조합이 없습니다.")

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
            except Exception as e:
                attempt += 1
                self.logger(f"❌ 전송/기타 오류: {repr(e)}")
                if attempt >= max_tries:
                    return False
                sleep_s = base_sleep * (2 ** (attempt - 1))
                self.logger(f"⏳ 재시도 {attempt}/{max_tries} ... {sleep_s:.1f}s")
                time.sleep(sleep_s)

# digits_only, is_int_string 등 유틸 함수는 config.py에서 import 하세요.


