# cellon/sellertool_excel.py
from __future__ import annotations

import shutil
from functools import lru_cache
from pathlib import Path
from typing import Iterable, Optional

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

import json
from functools import lru_cache
from pathlib import Path

from .config import (
    COUPANG_UPLOAD_FORM_DIR,
    UPLOAD_READY_DIR,
    SELLERTOOL_SHEET_NAME,
    COUPANG_UPLOAD_INDEX_JSON,
)
from .core.product import Product


# =========================
# 유틸
# =========================

def _safe_str(v) -> str:
    if v is None:
        return ""
    try:
        return str(v)
    except Exception:
        return ""


def _normalize_category_text(text: str) -> str:
    """
    템플릿 파일명과 category_path 를 최대한 유연하게 매칭하기 위한 정규화.

    - 공백 제거
    - 한글, 숫자, 영문만 남김
    - '/', ':', '>' 등을 모두 '>' 로 통합
    - 유니코드 정규화 적용
    """
    if not text:
        return ""

    import unicodedata
    t = unicodedata.normalize("NFKC", text)

    # 구분자 통합
    t = t.replace("/", ">").replace(":", ">")

    # 소문자
    t = t.lower()

    # 공백 제거
    t = t.replace(" ", "")

    # 한글/영문/숫자/구분자만 남김
    allowed = []
    for ch in t:
        if ch.isalnum() or ch == ">":
            allowed.append(ch)
    t = "".join(allowed)

    return t



# =========================
# 1) 템플릿 인덱스 (파일명 → Path)
# =========================

@lru_cache(maxsize=1)
def _build_template_index() -> dict[str, Path]:
    """
    쿠팡 업로드 템플릿 인덱스를 JSON 에서 읽어온다.

    - JSON 이 없으면: 카테고리 분석(또는 build_coupang_upload_index.py)을
      먼저 실행하라는 에러를 던진다.
    - JSON 포맷:
      {
        "templates": [
          { "key": "주방용품>취사도구", "relative_path": "주방용품/..." },
          ...
        ]
      }
    """
    if not COUPANG_UPLOAD_INDEX_JSON.exists():
        raise RuntimeError(
            f"쿠팡 업로드 템플릿 인덱스 JSON이 없습니다: {COUPANG_UPLOAD_INDEX_JSON}\n"
            "먼저 '카테고리 분석' 또는 build_coupang_upload_index.py 를 실행해서 "
            "인덱스를 생성해 주세요."
        )

    with COUPANG_UPLOAD_INDEX_JSON.open("r", encoding="utf-8") as f:
        data = json.load(f)

    templates = data.get("templates", [])
    if not templates:
        raise RuntimeError(
            f"쿠팡 업로드 템플릿 인덱스가 비어 있습니다: {COUPANG_UPLOAD_INDEX_JSON}\n"
            "카테고리 분석을 다시 실행해 주세요."
        )

    index: dict[str, Path] = {}
    root = COUPANG_UPLOAD_FORM_DIR

    for item in templates:
        key = item["key"]
        rel = Path(item["relative_path"])
        index[key] = (root / rel).resolve()

    return index


def find_template_for_category_path(category_path: str) -> Path:
    """
    쿠팡 카테고리 경로 (예: '주방용품>취사도구>냄비>양수냄비')
    에 가장 잘 매칭되는 템플릿 파일을 찾는다.

    전략:
      1) category_path 를 '>' 기준으로 자른 뒤
         가장 긴 prefix 부터 줄여가며(4뎁스 → 3뎁스 → 2뎁스 → 1뎁스)
         템플릿 key 와 매칭을 시도한다.
      2) prefix 와 템플릿 key 를 공백 제거한 문자열로 비교:
         - key_norm == prefix_norm
         - 또는 key_norm 이 prefix_norm 안에 포함
         - 또는 prefix_norm 이 key_norm 안에 포함
      3) 그래도 없으면, 전체 category_path 기준으로
         "가장 긴 부분 문자열로 겹치는" 기존 방식으로 한 번 더 시도.
    """
    index = _build_template_index()
    if not category_path:
        raise KeyError("카테고리 경로가 비어 있습니다.")

    parts = [p.strip() for p in category_path.split(">") if p.strip()]
    if not parts:
        raise KeyError(f"파싱할 수 없는 카테고리 경로입니다: {category_path}")

    # 미리 정규화된 key 캐시
    norm_index: dict[str, str] = {
        key: _normalize_category_text(key) for key in index.keys()
    }

    # 1) depth 를 줄여가며 prefix 매칭 시도
    for depth in range(len(parts), 0, -1):
        prefix = ">".join(parts[:depth])
        prefix_norm = _normalize_category_text(prefix)

        # 1-1) exact match 우선
        exact_matches: list[str] = [
            key for key, key_norm in norm_index.items()
            if key_norm == prefix_norm
        ]
        if exact_matches:
            # 가장 긴 key (뎁스 많은 것) 선택
            best_key = max(exact_matches, key=lambda k: len(norm_index[k]))
            return index[best_key]

        # 1-2) 포함 관계 매칭 (양방향)
        candidates: list[str] = [
            key for key, key_norm in norm_index.items()
            if key_norm in prefix_norm or prefix_norm in key_norm
        ]
        if candidates:
            # prefix 와 가장 비슷하게(문자열 길이 가장 긴 것) 선택
            best_key = max(candidates, key=lambda k: len(norm_index[k]))
            return index[best_key]

    # 2) 그래도 못 찾으면: 전체 category_path 기준으로
    #    "가장 긴 부분 문자열로 겹치는" 템플릿을 선택 (기존 fallback)
    category_norm = _normalize_category_text(category_path)
    best_key = None
    best_len = -1

    for key, key_norm in norm_index.items():
        if key_norm and (key_norm in category_norm or category_norm in key_norm):
            if len(key_norm) > best_len:
                best_key = key
                best_len = len(key_norm)

    if best_key is not None:
        return index[best_key]

    # 3) 완전히 매칭 실패했을 때
    available = ", ".join(sorted(index.keys()))
    raise KeyError(
        f"카테고리 경로에 맞는 템플릿을 찾지 못했습니다: {category_path} "
        f"(사용 가능한 템플릿 key: {available})"
    )

# =========================
# 2) data 시트 헬퍼
# =========================

def _get_header_col(ws: Worksheet, header_text: str) -> Optional[int]:
    """2행(헤더행)에서 header_text 와 정확히 일치하는 컬럼 index 를 찾는다."""
    for cell in ws[2]:
        if _safe_str(cell.value) == header_text:
            return cell.column
    return None


def _find_column_contains(ws: Worksheet, keyword: str) -> Optional[int]:
    """2행(헤더행)에서 keyword 를 포함하는 컬럼 index 를 찾는다."""
    for cell in ws[2]:
        val = _safe_str(cell.value)
        if keyword in val:
            return cell.column
    return None


def _find_first_empty_row(ws: Worksheet, start_row: int = 5) -> int:
    """
    A열 기준으로 첫 번째 빈 행을 찾는다.
    (A열에 카테고리 ID가 없으면 '입력 가능' 행이라고 가정)
    """
    row = start_row
    while True:
        if not _safe_str(ws.cell(row=row, column=1).value):
            return row
        row += 1


def _pick_template_row(
    ws: Worksheet,
    category_id: str,
    category_path: str,
) -> int:
    """
    - A열에 '[category_id]' 가 들어있는 행들을 후보로 모으고
    - 그 중 CK열('상품고시정보 카테고리') 값이
      카테고리 최상단(예: '주방용품')과 맞는 행을 우선 선택.
    """
    id_token = f"[{category_id}]"
    candidates: list[int] = []

    max_row = ws.max_row
    for r in range(5, max_row + 1):
        val = _safe_str(ws.cell(row=r, column=1).value)
        if id_token in val:
            candidates.append(r)

    if not candidates:
        raise RuntimeError(f"data 시트에서 category_id={category_id} 행을 찾지 못했습니다.")

    # CK열 (상품고시정보 카테고리) 인덱스 찾기
    ck_col = _find_column_contains(ws, "상품고시정보 카테고리")
    top_level = _safe_str(category_path).split(">")[0].strip()

    if ck_col and top_level:
        for r in candidates:
            ck_val = _safe_str(ws.cell(row=r, column=ck_col).value)
            if top_level in ck_val:
                return r

    # 적절한 CK 매칭이 없으면 그냥 첫 번째 후보 사용
    return candidates[0]


def _copy_row(ws: Worksheet, src_row: int, dst_row: int, max_col: int = 120) -> None:
    """
    src_row → dst_row 로 '값'만 복사.
    (스타일/데이터유효성까지 필요해지면 나중에 확장)
    """
    for col in range(1, max_col + 1):
        src = ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)
        dst.value = src.value


def _fill_product_data(
    ws: Worksheet,
    row: int,
    *,
    product: Product,
    price: Optional[int] = None,
    search_keywords: Optional[Iterable[str]] = None,
) -> None:
    """
    템플릿 복사된 행(row)에 실제 상품 데이터를 채워 넣는다.
    - 등록상품명
    - 판매가격
    - 검색어
    등 기본적인 것만 우선 채우고, 나머지는 나중에 확장.
    """
    # 1) 헤더 컬럼 인덱스 찾기
    col_name = _get_header_col(ws, "등록상품명")
    col_price = _get_header_col(ws, "판매가격")
    col_search = _get_header_col(ws, "검색어")

    # 2) 상품명
    if col_name:
        ws.cell(row=row, column=col_name).value = product.display_name

    # 3) 판매가격
    if col_price is not None and price is not None:
        try:
            ws.cell(row=row, column=col_price).value = int(price)
        except Exception:
            ws.cell(row=row, column=col_price).value = _safe_str(price)

    # 4) 검색어
    if col_search is not None and search_keywords:
        joined = ", ".join([_safe_str(k) for k in search_keywords if _safe_str(k)])
        ws.cell(row=row, column=col_search).value = joined


# =========================
# 3) 퍼블릭 API
# =========================

def prepare_and_fill_sellertool(
    *,
    product: Product,
    coupang_category_id: str,
    coupang_category_path: str,
    price: Optional[int] = None,
    search_keywords: Optional[Iterable[str]] = None,
) -> Path:
    """
    1) coupang_upload_form 폴더에서 카테고리에 맞는 템플릿 엑셀을 찾고
    2) upload_ready 폴더로 '원래 파일명 그대로' 복사 (이미 있으면 재사용)
    3) data 시트에서 category_id/경로에 맞는 템플릿 행을 찾아
       첫 빈 행으로 복사하고
    4) 그 행에 product/price/search_keywords 를 채워 넣는다.

    최종적으로 수정된 upload_ready 안의 파일 Path 를 반환.
    """
       
    if not COUPANG_UPLOAD_INDEX_JSON.exists():
        raise RuntimeError(
            f"쿠팡 업로드 템플릿 인덱스 JSON이 없습니다: {COUPANG_UPLOAD_INDEX_JSON}\n"
            "먼저 '카테고리 분석' 또는 build_coupang_upload_index.py 를 실행해 주세요."
        )
    # ---- 디버그 로그 엑셀파일이 아닌 json 에서 읽는지 확인 ----
    print("[DEBUG] template index JSON exists. loading:", COUPANG_UPLOAD_INDEX_JSON)
    
    with COUPANG_UPLOAD_INDEX_JSON.open("r", encoding="utf-8") as f:
        data = json.load(f)
    # ---- 디버그 로그 엑셀파일이 아닌 json 에서 읽는지 확인 ----
    
        
    # ---- 1) 템플릿 선택 ----
    template_path = find_template_for_category_path(coupang_category_path)

    # ---- 2) upload_ready 폴더로 '원래 파일명' 그대로 복사 ----
    UPLOAD_READY_DIR.mkdir(parents=True, exist_ok=True)
    dest_path = UPLOAD_READY_DIR / template_path.name

    # 같은 템플릿을 여러 번 쓰는 경우:
    # 이미 dest_path 가 있으면 복사하지 않고 기존 파일에 행만 추가
    if not dest_path.exists():
        shutil.copy2(template_path, dest_path)

    # ---- 3) 엑셀 열기 ----
    wb = load_workbook(dest_path, keep_vba=True, data_only=False)

    if SELLERTOOL_SHEET_NAME not in wb.sheetnames:
        raise RuntimeError(
            f"시트 '{SELLERTOOL_SHEET_NAME}' 를 찾지 못했습니다. "
            f"파일: {dest_path}"
        )

    ws = wb[SELLERTOOL_SHEET_NAME]

    # ---- 4) 템플릿 행 선택 (category_id + top-level 카테고리) ----
    src_row = _pick_template_row(ws, coupang_category_id, coupang_category_path)

    # ---- 5) 첫 번째 빈 입력 행 찾기 (A열 기준) ----
    dst_row = _find_first_empty_row(ws, start_row=5)

    # ---- 6) 템플릿 행 복사 후 상품 데이터 덮어쓰기 ----
    _copy_row(ws, src_row, dst_row, max_col=120)
    _fill_product_data(
        ws,
        dst_row,
        product=product,
        price=price,
        search_keywords=search_keywords,
    )

    # ---- 7) 저장 ----
    wb.save(dest_path)

    return dest_path
