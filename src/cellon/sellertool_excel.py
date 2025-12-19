# cellon/sellertool_excel.py
from __future__ import annotations

import json
import shutil
from copy import copy
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import Iterable, Optional

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .config import (
    COUPANG_UPLOAD_FORM_DIR,
    UPLOAD_READY_DIR,
    SELLERTOOL_SHEET_NAME,
    COUPANG_UPLOAD_INDEX_JSON,
)
from .core.product import Product

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from typing import Dict, Tuple

_WB_CACHE: Dict[str, Tuple[Workbook, float]] = {}
# key: str(dest_path), value: (workbook, last_mtime)

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
# 1) 템플릿 인덱스 (파일명 → 절대경로 폴더 내 재귀검색)
# =========================

@lru_cache(maxsize=1)
def _build_template_index() -> dict[str, Path]:
    """
    쿠팡 업로드 템플릿 인덱스를 생성한다.

    기본 정책:
    - (B) JSON 인덱스가 있으면: JSON 을 신뢰하고 빠르게 로드한다. (추천)
    - (A) JSON 인덱스가 없거나/깨졌거나/비어 있으면: rglob() 재귀 탐색으로 백업한다.

    - JSON 포맷:
      {
        "templates": [
          { "key": "주방용품>취사도구", "relative_path": "주방용품/..." },
          ...
        ]
      }
    """
    root = COUPANG_UPLOAD_FORM_DIR

    # -------------------------
    # (B) JSON 인덱스 우선 로드
    # -------------------------
    if COUPANG_UPLOAD_INDEX_JSON.exists():
        try:
            with COUPANG_UPLOAD_INDEX_JSON.open("r", encoding="utf-8") as f:
                data = json.load(f)

            templates = data.get("templates", [])
            if not templates:
                # JSON은 있으나 내용이 비어있음 → 기존처럼 강하게 안내 + (A) 백업 시도
                print(
                    f"[WARN] 쿠팡 업로드 템플릿 인덱스가 비어 있습니다: {COUPANG_UPLOAD_INDEX_JSON}\n"
                    "카테고리 분석을 다시 실행하는 것을 권장합니다. (백업 탐색을 시도합니다)"
                )
            else:
                index: dict[str, Path] = {}
                missing_count = 0

                for item in templates:
                    key = item.get("key")
                    rel_raw = item.get("relative_path")

                    if not key or not rel_raw:
                        # 포맷 이상 항목은 건너뛰되 경고만 남김
                        print(f"[WARN] 잘못된 템플릿 인덱스 항목을 건너뜁니다: {item}")
                        continue

                    rel = Path(rel_raw)
                    abs_path = (root / rel).resolve()
                    index[key] = abs_path

                    # 파일이 실제로 없으면 카운트만 하고 계속 (나중에 백업 여부 결정)
                    if not abs_path.exists():
                        missing_count += 1

                # JSON 기반 인덱스가 실질적으로 유효하면 그대로 사용
                if index and missing_count == 0:
                    return index

                # 일부/전체 경로가 깨진 경우 → 경고 후 (A) 백업 탐색
                if index and missing_count > 0:
                    print(
                        "[WARN] 쿠팡 템플릿 인덱스 JSON은 있으나, 실제 파일 경로가 누락된 항목이 있습니다.\n"
                        f"- 누락 항목 수: {missing_count}\n"
                        f"- JSON 경로: {COUPANG_UPLOAD_INDEX_JSON}\n"
                        "해결:\n"
                        "1) 템플릿 파일 이동/삭제 여부를 확인하거나\n"
                        "2) build_coupang_upload_index.py 를 다시 실행해 인덱스를 재생성하세요.\n"
                        "우선 백업 탐색(rglob)을 시도합니다."
                    )

                # JSON이 있긴 하나 결과가 비어 있거나 경로가 깨짐 → 백업으로 진행
        except Exception as e:
            # JSON 파손/인코딩 문제 등 → 기존처럼 강하게 안내 + (A) 백업 시도
            print(
                f"[WARN] 쿠팡 업로드 템플릿 인덱스 JSON 로드에 실패했습니다: {COUPANG_UPLOAD_INDEX_JSON}\n"
                f"- 원인: {repr(e)}\n"
                "해결:\n"
                "1) 카테고리 분석(또는 build_coupang_upload_index.py)을 다시 실행하거나\n"
                "2) JSON 파일이 정상인지 확인하세요.\n"
                "우선 백업 탐색(rglob)을 시도합니다."
            )

    else:
        # 기존 안전장치 메시지 톤을 유지하되, 즉시 종료하지 않고 백업을 시도
        print(
            f"[WARN] 쿠팡 업로드 템플릿 인덱스 JSON이 없습니다: {COUPANG_UPLOAD_INDEX_JSON}\n"
            "먼저 '카테고리 분석' 또는 build_coupang_upload_index.py 를 실행해서 "
            "인덱스를 생성하는 것을 권장합니다.\n"
            "우선 백업 탐색(rglob)을 시도합니다."
        )

    # -------------------------
    # (A) 백업: rglob 재귀 탐색
    # -------------------------
    index: dict[str, Path] = {}
    for path in root.rglob("sellertool_upload_*.xlsm"):
        key = path.stem.replace("sellertool_upload_", "")
        index[key] = path.resolve()

    if index:
        return index

    # -------------------------
    # 최종 실패: 기존과 유사한 강한 에러
    # -------------------------
    raise RuntimeError(
        f"쿠팡 업로드 폼 템플릿을 찾지 못했습니다: {root}\n"
        "확인:\n"
        "1) coupang_upload_form 내의 쿠팡 템플릿 구조 아래에 sellertool_upload_*.xlsm 이 존재하는지\n"
        f"2) 인덱스 JSON({COUPANG_UPLOAD_INDEX_JSON})이 정상인지\n"
        "해결:\n"
        "1) '카테고리 분석' 또는 build_coupang_upload_index.py 로 인덱스를 생성하고\n"
        "2) 템플릿 파일들이 올바른 위치에 있는지 확인해 주세요."
    )




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

def _get_target_insertion_row(ws: Worksheet) -> int:
    """
    [새로 추가된 함수]
    파일마다 다른 양식 구간(Template Source)의 끝을 찾아내고,
    그 아래에 구분자(Divider)를 넣은 뒤 실제 데이터가 들어갈 행 번호를 반환합니다.
    """
    # 1. 시트에서 실제로 데이터가 있는 마지막 행 찾기 (A열 기준 역순 탐색)
    # 엑셀 파일마다 10행일수도, 128행일수도, 300행일수도 있는 '양식 끝'을 동적으로 찾습니다.
    last_template_row = 0
    for r in range(ws.max_row, 0, -1):
        if _safe_str(ws.cell(row=r, column=1).value):
            last_template_row = r
            break
            
    # 2. 이미 구분자("---")가 포함된 행이 있는지 확인 (이미 상품이 하나 이상 추가된 경우)
    divider_found_row = 0
    for r in range(1, ws.max_row + 1):
        val = _safe_str(ws.cell(row=r, column=1).value)
        if "------------------" in val:
            divider_found_row = r
            break

    # 3. 상황에 따른 목적지 행(Destination Row) 결정
    if divider_found_row > 0:
        # CASE A: 이미 구분자가 있음 -> 구분자 아래에서 첫 번째 빈 행을 찾아 이어서 작성
        curr_row = divider_found_row + 1
        while True:
            if not _safe_str(ws.cell(row=curr_row, column=1).value):
                return curr_row
            curr_row += 1
    else:
        # CASE B: 구분자가 없음 (최초 작성) -> 양식 끝 바로 다음 행에 구분자 삽입
        divider_row = last_template_row + 1
        ws.cell(row=divider_row, column=1).value = "------------------ 여기서부터 크롤링 데이터 등록 ------------------"
        
        # 가독성을 위해 노란색 배경색(PatternFill) 추가 (선택사항)
        from openpyxl.styles import PatternFill
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws.cell(row=divider_row, column=1).fill = yellow_fill
        
        # 데이터는 구분자 바로 다음 행(양식끝 + 2)부터 시작
        return divider_row + 1



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


def _copy_row_full(ws: Worksheet, src_row: int, dst_row: int, max_col: int = 120) -> None:
    """
    [수정] src_row → dst_row 로 값, 스타일, 데이터 유효성(드롭다운)까지 모두 복사.
    단순 값 복사에서 '스타일(색상, 테두리, 수식 등)'까지 복사하도록 변경.
    쿠팡 엑셀은 양식 서식이 중요하므로 copy 모듈을 사용합니다.
    """
    # 1. 스타일 및 값 복사
    for col in range(1, max_col + 1):
        src_cell = ws.cell(row=src_row, column=col)
        dst_cell = ws.cell(row=dst_row, column=col)
        dst_cell.value = src_cell.value
        
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.border = copy(src_cell.border)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.number_format = copy(src_cell.number_format)
            dst_cell.alignment = copy(src_cell.alignment)

    # 2. 데이터 유효성(드롭다운) 복사
    for dv in list(ws.data_validations.dataValidation):
        ranges_snapshot = list(dv.sqref.ranges)
        for range_obj in ranges_snapshot:
            if range_obj.min_row <= src_row <= range_obj.max_row:
                for col in range(range_obj.min_col, range_obj.max_col + 1):
                    addr = f"{get_column_letter(col)}{dst_row}"
                    dv.add(addr)

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



def _get_cached_workbook(xlsm_path: Path) -> Workbook:
    """
    같은 xlsm을 반복해서 열지 않기 위한 간단 캐시.
    - 파일 mtime이 바뀌면(외부 수정/복사 등) 캐시 무효화 후 재로딩
    """
    key = str(xlsm_path)
    mtime = xlsm_path.stat().st_mtime

    cached = _WB_CACHE.get(key)
    if cached:
        wb, cached_mtime = cached
        if cached_mtime == mtime:
            return wb

    wb = load_workbook(xlsm_path, keep_vba=True)
    _WB_CACHE[key] = (wb, mtime)
    return wb


def _save_cached_workbook(xlsm_path: Path, wb: Workbook) -> None:
    """
    저장 후 mtime 갱신(캐시 유지)
    """
    wb.save(xlsm_path)
    _WB_CACHE[str(xlsm_path)] = (wb, xlsm_path.stat().st_mtime)


def prepare_sellertool_workbook_copy(
    template_xlsm_path: Path,
    out_dir: Path,
    output_name: str | None = None,
    add_date_subdir: bool = False,
) -> Path:
    """
    ✅ output_name이 None이면 '템플릿 파일명 그대로' 복사/재사용
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if add_date_subdir:
        out_dir = out_dir / datetime.now().strftime("%Y-%m-%d")
        out_dir.mkdir(parents=True, exist_ok=True)

    if output_name:
        dest_path = out_dir / output_name
    else:
        dest_path = out_dir / Path(template_xlsm_path).name  # ✅ 파일명 유지

    # ✅ 이미 upload_ready에 있으면 복사하지 않음(재사용)
    if not dest_path.exists():
        shutil.copy2(template_xlsm_path, dest_path)

    return dest_path


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
    #for test
    
    print("[DEBUG] prepare_and_fill_sellertool called")
    
    # ---- 디버그 로그: 템플릿 인덱스 JSON 상태 확인 ----
    # 정책:
    # - JSON 이 있으면 (B) 인덱스 기반으로 빠르게 찾는다.
    # - JSON 이 없거나/깨졌으면 (A) rglob 백업 탐색으로 찾는다.
    if COUPANG_UPLOAD_INDEX_JSON.exists():
        print("[DEBUG] template index JSON exists:", COUPANG_UPLOAD_INDEX_JSON)
    else:
        print(
            "[WARN] 쿠팡 업로드 템플릿 인덱스 JSON이 없습니다.\n"
            f"- JSON 경로: {COUPANG_UPLOAD_INDEX_JSON}\n"
            "권장:\n"
            "1) '카테고리 분석' 또는 build_coupang_upload_index.py 를 실행해 "
            "인덱스를 생성해 주세요.\n"
            "우선 백업 탐색(rglob)으로 템플릿 선택을 시도합니다."
        )

    
        
    # ---- 1) 템플릿 선택 ----
    template_path = find_template_for_category_path(coupang_category_path)

    # ---- 2) upload_ready 폴더로 '원래 파일명' 그대로 복사 ----
    UPLOAD_READY_DIR.mkdir(parents=True, exist_ok=True)
    dest_name = template_path.name
    dest_path = UPLOAD_READY_DIR / dest_name

    # 같은 템플릿을 여러 번 쓰는 경우:
    # 이미 dest_path 가 있으면 복사하지 않고 기존 파일에 행만 추가
    if not dest_path.exists():
        shutil.copy2(template_path, dest_path)

    # ---- 3) 엑셀 열기 ----
    wb = _get_cached_workbook(dest_path)

    if SELLERTOOL_SHEET_NAME not in wb.sheetnames:
        raise RuntimeError(
            f"시트 '{SELLERTOOL_SHEET_NAME}' 를 찾지 못했습니다. "
            f"파일: {dest_path}"
        )

    ws = wb[SELLERTOOL_SHEET_NAME]

    # ---- 4) 템플릿 행 선택 (category_id + top-level 카테고리) ----
    src_row = _pick_template_row(ws, coupang_category_id, coupang_category_path)

    # ---- 5) 첫 번째 빈 입력 행 찾기 (A열 기준) ----
    dst_row = _get_target_insertion_row(ws)

    # ---- 6) 템플릿 행 복사 후 상품 데이터 덮어쓰기 ----
    _copy_row_full(ws, src_row, dst_row, max_col=120)
    _fill_product_data(
        ws,
        dst_row,
        product=product,
        price=price,
        search_keywords=search_keywords,
    )

    # ---- 7) 저장 ----
    _save_cached_workbook(dest_path, wb)
    
    
    print("[DEBUG] template_path =", template_path)
    print("[DEBUG] template_path.name =", template_path.name)
    print("[DEBUG] dest_path =", dest_path)
    print("[DEBUG] dest exists? =", dest_path.exists())

    return dest_path


