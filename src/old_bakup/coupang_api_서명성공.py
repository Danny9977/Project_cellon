# ================== 쿠팡 OpenAPI v2 서명 생성 및 요청 예제 - 성공 - ==================
import os
import hmac
import ssl
import json
import hashlib
import urllib.parse
import urllib.request
from pathlib import Path
from datetime import datetime, timezone

# ====== 1) JSON 키 로드 (snake_case) ======
COUPANG_KEYS_JSON = Path("/Users/jeehoonkim/Desktop/Python_Project/api/coupang_api/coupang_keys.json")
with open(COUPANG_KEYS_JSON, "r") as f:
    keys = json.load(f)

vendor_id = keys.get("vendor_id")
access_key = keys.get("access_key")
secret_key = keys.get("secret_key")

if not all([vendor_id, access_key, secret_key]):
    raise ValueError("❌ coupang_keys.json에 vendor_id / access_key / secret_key 중 누락이 있습니다.")

# ====== 2) 서명 시간: UTC, YYMMDDTHHMMSSZ (하이픈/콜론 없음) ======
# 예: 251111T110106Z
dt_utc = datetime.now(timezone.utc)
signed_date = dt_utc.strftime("%y%m%dT%H%M%SZ")
method = "GET"

# ====== 3) 경로/쿼리 (서명용 query == URL query 동일) ======
path = f"/v2/providers/openapi/apis/api/v4/vendors/{vendor_id}/returnRequests"
query_params = {
    "createdAtFrom": "2018-08-08",
    "createdAtTo": "2018-08-09",
    "status": "UC"
}
query = urllib.parse.urlencode(query_params)

# ====== 4) HMAC 메시지 구성 및 서명 생성 ======
# message = signed-date + method + path + query  (구분자 없음, '?' 없음)
message = f"{signed_date}{method}{path}{query}"
signature = hmac.new(
    secret_key.encode("utf-8"),
    message.encode("utf-8"),
    hashlib.sha256
).hexdigest()

authorization = (
    f"CEA algorithm=HmacSHA256, access-key={access_key}, "
    f"signed-date={signed_date}, signature={signature}"
)

# ====== 5) 요청 ======
url = f"https://api-gateway.coupang.com{path}?{query}"
req = urllib.request.Request(url)
req.add_header("Content-Type", "application/json;charset=UTF-8")
req.add_header("Authorization", authorization)
req.get_method = lambda: method

# (테스트용) SSL 검증 비활성화 — 운영에서는 제거 권장
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

print("Request URL:", req.get_full_url())
print("Signed-Date:", signed_date)
print("Message for HMAC:", message)            # 디버깅 후 제거 권장
print("Authorization(head):", authorization[:70] + "...")

try:
    resp = urllib.request.urlopen(req, context=ctx)
    body = resp.read().decode(resp.headers.get_content_charset() or "utf-8")
    print("✅ [SUCCESS]", resp.getcode())
    print(body)
except urllib.error.HTTPError as e:
    print(f"❌ HTTP Error: {e.code} - {e.reason}")
    err_body = e.read().decode("utf-8", errors="replace")
    print(err_body)
except urllib.error.URLError as e:
    print(f"❌ URL Error: {e.reason}")
except Exception as e:
    print(f"❌ Unexpected Error: {e}")
