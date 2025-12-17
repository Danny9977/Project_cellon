# src/cellon/apis/coupang_client.py
from __future__ import annotations

import json
import hmac
import hashlib
import requests
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode
from functools import lru_cache

from ..config import COUPANG_KEYS_JSON, COUPANG_BASE_URL

# =========================
# 쿠팡 OpenAPI HMAC 서명 (성공 예제 기준)
# =========================


@dataclass(frozen=True)
class CoupangKeys:
    vendor_id: str
    access_key: str
    secret_key: str


@lru_cache(maxsize=4)
def load_coupang_keys(json_path: str | Path | None = None) -> CoupangKeys:
    """
    쿠팡 키를 coupang_keys.json에서 읽어 단일 객체로 반환.
    - 어디서 호출하든 동일한 방식/검증 로직을 사용하도록 '단일 함수'로 통합.
    """
    p = Path(json_path) if json_path else Path(COUPANG_KEYS_JSON)
    if not p.exists():
        raise FileNotFoundError(f"쿠팡 키 파일을 찾지 못했습니다: {p}")

    with open(p, "r", encoding="utf-8") as f:
        data = json.load(f)

    vendor_id = (data.get("vendor_id") or "").strip()
    access_key = (data.get("access_key") or "").strip()
    secret_key = (data.get("secret_key") or "").strip()

    if not vendor_id:
        raise ValueError("coupang_keys.json의 vendor_id가 비어 있습니다.")
    if not access_key:
        raise ValueError("coupang_keys.json의 access_key가 비어 있습니다.")
    if not secret_key:
        raise ValueError("coupang_keys.json의 secret_key가 비어 있습니다.")

    return CoupangKeys(vendor_id=vendor_id, access_key=access_key, secret_key=secret_key)


def cp_build_query(params: dict | None) -> str:
    if not params:
        return ""
    return urlencode(params, doseq=True)


def cp_signed_headers_v2(
    method: str,
    path: str,
    sign_query: str,
    access_key: str,
    secret_key: str,
    *,
    signed_date: str | None = None,
    vendor_id: str | None = None,
) -> dict:
    if signed_date is None:
        signed_date = datetime.now(timezone.utc).strftime("%y%m%dT%H%M%SZ")

    message = f"{signed_date}{method.upper()}{path}{sign_query}"
    signature = hmac.new(
        secret_key.encode("utf-8"),
        message.encode("utf-8"),
        hashlib.sha256,
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


def cp_request(method: str, path: str, params: dict | None = None, *, keys: CoupangKeys | None = None) -> dict:
    """
    쿠팡 OpenAPI 요청 (서명/헤더/키 로딩을 여기서 일원화)
    """
    if keys is None:
        keys = load_coupang_keys()

    url_query = cp_build_query(params)
    url = f"{COUPANG_BASE_URL}{path}" + (f"?{url_query}" if url_query else "")

    headers = cp_signed_headers_v2(
        method,
        path,
        url_query,
        keys.access_key,
        keys.secret_key,
        vendor_id=keys.vendor_id,
    )

    resp = requests.request(method=method, url=url, headers=headers, timeout=30)
    try:
        resp.raise_for_status()
    except requests.HTTPError as e:
        body = (resp.text or "")[:1000]
        raise requests.HTTPError(
            f"{resp.status_code} {resp.reason}\nurl={url}\nresp_body={body}",
            response=resp,
            request=resp.request,
        ) from e

    return resp.json()


def build_ordersheets_params(date_from_utc: datetime, date_to_utc: datetime, status: str, max_per_page: int = 50):
    d_from = date_from_utc.strftime("%Y-%m-%d")
    d_to = date_to_utc.strftime("%Y-%m-%d")

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


def try_ordersheets_with_variants(path: str, param_variants: list[dict], *, keys: CoupangKeys | None = None) -> dict:
    last_err = None
    for params in param_variants:
        try:
            return cp_request("GET", path, params, keys=keys)
        except requests.HTTPError as e:
            resp = getattr(e, "response", None)
            status = getattr(resp, "status_code", None)
            body = ""
            try:
                body = (resp.text or "")[:500]
            except Exception:
                pass

            # 날짜 포맷 관련 400이면 fallback 파라미터로 계속 시도
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
