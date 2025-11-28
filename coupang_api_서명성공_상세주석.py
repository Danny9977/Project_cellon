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
COUPANG_KEYS_JSON = Path("/Users/jeehoonkim/Desktop/Python_Project/api/coupang_api/coupang_keys.json")  # API 키 파일 경로 지정
with open(COUPANG_KEYS_JSON, "r") as f:  # JSON 파일 열기
    keys = json.load(f)  # JSON 데이터 파싱

vendor_id = keys.get("vendor_id")         # 벤더 ID 추출
access_key = keys.get("access_key")       # 액세스 키 추출
secret_key = keys.get("secret_key")       # 시크릿 키 추출

if not all([vendor_id, access_key, secret_key]):  # 키 값이 모두 있는지 확인
    raise ValueError("❌ coupang_keys.json에 vendor_id / access_key / secret_key 중 누락이 있습니다.")

# ====== 2) 서명 시간: UTC, YYMMDDTHHMMSSZ (하이픈/콜론 없음) ======
# 예: 251111T110106Z
dt_utc = datetime.now(timezone.utc)                   # 현재 UTC 시간 가져오기
signed_date = dt_utc.strftime("%y%m%dT%H%M%SZ")       # 서명용 날짜 포맷으로 변환
method = "GET"                                        # HTTP 메서드 지정

# ====== 3) 경로/쿼리 (서명용 query == URL query 동일) ======
path = f"/v2/providers/openapi/apis/api/v4/vendors/{vendor_id}/returnRequests"  # API 엔드포인트 경로
query_params = {                                         # 쿼리 파라미터 정의
    "createdAtFrom": "2018-08-08",
    "createdAtTo": "2018-08-09",
    "status": "UC"
}
query = urllib.parse.urlencode(query_params)             # 쿼리 파라미터 URL 인코딩

# ====== 4) HMAC 메시지 구성 및 서명 생성 ======
# message = signed-date + method + path + query  (구분자 없음, '?' 없음)
message = f"{signed_date}{method}{path}{query}"         # 서명용 메시지 생성
signature = hmac.new(                                   # HMAC-SHA256 서명 생성
    secret_key.encode("utf-8"),
    message.encode("utf-8"),
    hashlib.sha256
).hexdigest()

authorization = (                                       # Authorization 헤더 포맷팅
    f"CEA algorithm=HmacSHA256, access-key={access_key}, "
    f"signed-date={signed_date}, signature={signature}"
)

# ====== 5) 요청 ======
url = f"https://api-gateway.coupang.com{path}?{query}"  # 최종 요청 URL 생성
req = urllib.request.Request(url)                       # Request 객체 생성
req.add_header("Content-Type", "application/json;charset=UTF-8")  # Content-Type 헤더 추가
req.add_header("Authorization", authorization)          # Authorization 헤더 추가
req.get_method = lambda: method                        # HTTP 메서드 지정 (GET)

# (테스트용) SSL 검증 비활성화 — 운영에서는 제거 권장
ctx = ssl.create_default_context()                      # SSL 컨텍스트 생성
ctx.check_hostname = False                             # 호스트네임 검증 비활성화
ctx.verify_mode = ssl.CERT_NONE                        # 인증서 검증 비활성화

print("Request URL:", req.get_full_url())              # 요청 URL 출력
print("Signed-Date:", signed_date)                     # 서명 날짜 출력
print("Message for HMAC:", message)                    # HMAC 메시지 출력 (디버깅용)
print("Authorization(head):", authorization[:70] + "...")  # Authorization 헤더 일부 출력

try:
    resp = urllib.request.urlopen(req, context=ctx)    # API 요청 전송
    body = resp.read().decode(resp.headers.get_content_charset() or "utf-8")  # 응답 본문 디코딩
    print("✅ [SUCCESS]", resp.getcode())               # 성공 코드 출력
    print(body)                                        # 응답 본문 출력
except urllib.error.HTTPError as e:                    # HTTP 에러 처리
    print(f"❌ HTTP Error: {e.code} - {e.reason}")
    err_body = e.read().decode("utf-8", errors="replace")
    print(err_body)
except urllib.error.URLError as e:                     # URL 에러 처리
    print(f"❌ URL Error: {e.reason}")
except Exception as e:                                 # 기타 예외 처리
    print(f"❌ Unexpected Error: {e}")