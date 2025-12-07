
# ==== 다른 코드에서 이 모듈을 사용하는 예시: ====
# from cellon.category_ai.category_loader import load_category_master
# df = load_category_master()
# =================================================


# cellon/category_ai/category_loader.py. asdfasdfasdf
import os  # 폴더/파일 목록을 다루는 모듈
import re  # 문자열에서 패턴 찾기 위한 모듈
import hashlib  # 문자열을 해시값으로 변환하는 모듈
from pathlib import Path  # 파일/폴더 경로를 쉽게 다루는 모듈
from typing import List, Callable, Optional  # 타입 힌트용

import pandas as pd  # 엑셀 등 표 데이터를 다루는 라이브러리

# ===== 경로 설정 =====

from pathlib import Path
from typing import List, Callable, Optional
import pandas as pd
import os, re, hashlib

from cellon.config import CATEGORY_EXCEL_DIR, CACHE_DIR as CONFIG_CACHE_DIR

BASE_DIR = Path(__file__).resolve().parent.parent.parent
CATEGORY_DIR = CATEGORY_EXCEL_DIR          # ← 엑셀 위치는 config에서 통일 관리

CACHE_DIR = CONFIG_CACHE_DIR               # ← cache 위치도 config와 동일하게
CACHE_DIR.mkdir(parents=True, exist_ok=True)

CACHE_FILES_DIR = CACHE_DIR / "category_files"
CACHE_FILES_DIR.mkdir(parents=True, exist_ok=True)

MASTER_CACHE_FILE = CACHE_DIR / "category_master.pkl"

def _file_key(path: Path) -> str:  # 파일 경로를 해시값(고유키)로 변환
    h = hashlib.sha1(str(path.resolve()).encode("utf-8")).hexdigest()  # 경로를 sha1 해시로 변환
    return h  # 해시값 반환

# ===== 1) 개별 엑셀에서 카테고리 추출 =====
def extract_categories_from_file(path: str) -> pd.DataFrame:  # 엑셀 1개에서 카테고리 추출
    df = pd.read_excel(path, sheet_name="data", header=None)  # 'data' 시트 전체 읽기
    col0 = df.iloc[:, 0]  # 첫 번째 컬럼만 추출

    cat_rows = []  # 카테고리 정보 저장 리스트
    for v in col0:  # 첫 컬럼의 모든 셀 반복
        if isinstance(v, str) and "[" in v and "]" in v:  # [숫자]가 포함된 문자열만 처리
            m = re.match(r"\[(\d+)\]\s*(.+)", v.strip())  # [숫자]와 경로 분리
            if not m:  # 패턴 안 맞으면 건너뜀
                continue
            cat_id = m.group(1)  # 카테고리ID 추출
            path_str = m.group(2)  # 경로 문자열 추출
            cat_rows.append((cat_id, path_str))  # 튜플로 리스트에 추가

    if not cat_rows:  # 카테고리 정보 없으면
        return pd.DataFrame(
            columns=["category_id", "category_path", "level1", "level2", "level3", "level4"]  # 빈 테이블 반환
        )

    cat_rows = list(dict.fromkeys(cat_rows))  # 중복 제거

    records = []  # 최종 테이블에 넣을 딕셔너리 리스트
    for cat_id, path_str in cat_rows:  # 추출한 카테고리 튜플 반복
        parts = [p.strip() for p in path_str.split(">")]  # '>'로 경로 분리
        parts = parts + [""] * (4 - len(parts))  # 단계가 4개보다 적으면 빈칸 채움
        level1, level2, level3, level4 = parts[:4]  # 최대 4단계까지만 사용
        records.append(
            {
                "category_id": cat_id,  # 카테고리ID
                "category_path": path_str,  # 전체 경로 문자열
                "level1": level1,  # 1단계
                "level2": level2,  # 2단계
                "level3": level3,  # 3단계
                "level4": level4,  # 4단계
            }
        )

    return pd.DataFrame(records)  # 모든 카테고리 정보를 테이블로 반환

# ===== 2) 증분 방식 카테고리 마스터 생성 =====
def get_category_master(
    category_dir: Path = CATEGORY_DIR,
    progress_cb: Optional[Callable[[int, str], None]] = None,
) -> pd.DataFrame:
    """
    여러 엑셀 파일을 읽어서, 캐시가 있으면 캐시 사용, 없으면 새로 분석
    모든 결과를 합쳐서 마스터 테이블로 만듦
    진행상황을 콜백 함수로 알릴 수 있음
    """
    excel_files: List[Path] = []  # 엑셀 파일 경로 리스트
    for fname in os.listdir(category_dir):  # 폴더 내 파일 반복
        if not fname.lower().endswith(".xlsx"):  # 엑셀 파일만 처리
            continue
        if fname.startswith("~$"):  # 임시파일은 건너뜀
            continue
        excel_files.append(category_dir / fname)  # 전체 경로 리스트에 추가

    total = len(excel_files)  # 전체 파일 개수
    if total == 0:  # 파일 없으면
        if progress_cb:
            progress_cb(100, "카테고리 엑셀 파일이 없습니다.")  # 콜백으로 알림
        return pd.DataFrame(
            columns=["category_id", "category_path", "level1", "level2", "level3", "level4"]  # 빈 테이블 반환
        )

    if progress_cb:
        progress_cb(0, f"카테고리 엑셀 {total}개 분석 시작...")  # 시작 메시지

    per_file_dfs: List[pd.DataFrame] = []  # 각 파일별 결과 리스트
    processed = 0  # 처리한 파일 개수
    last_reported_percent = -5  # 진행률 보고용 변수

    for excel_path in excel_files:  # 모든 엑셀 파일 반복
        processed += 1  # 처리 파일 개수 증가
        key = _file_key(excel_path)  # 파일 경로 해시값 생성
        cache_pkl = CACHE_FILES_DIR / f"{key}.pkl"  # 캐시 파일 경로

        excel_mtime = excel_path.stat().st_mtime  # 엑셀 파일 수정시간
        use_cache = False  # 캐시 사용 여부

        if cache_pkl.exists():  # 캐시 파일이 있으면
            cache_mtime = cache_pkl.stat().st_mtime  # 캐시 파일 수정시간
            if cache_mtime >= excel_mtime:  # 캐시가 최신이면
                use_cache = True  # 캐시 사용

        if use_cache:  # 캐시 사용 시
            df = pd.read_pickle(cache_pkl)  # 캐시에서 데이터 읽기
            status = "캐시 사용"  # 상태 메시지
        else:  # 캐시가 없거나 오래됐으면
            df = extract_categories_from_file(str(excel_path))  # 엑셀에서 데이터 추출
            df.to_pickle(cache_pkl)  # 캐시 파일로 저장
            status = "재분석 및 캐시 갱신"  # 상태 메시지

        per_file_dfs.append(df)  # 결과 리스트에 추가

        # 진행률 계산 및 5% 단위로만 로그
        if progress_cb and total > 0:
            percent = int(processed / total * 100)  # 진행률 계산
            if percent - last_reported_percent >= 5:  # 5% 이상 변하면
                last_reported_percent = percent  # 마지막 보고값 갱신
                msg = f"[{percent}%] {excel_path.name} 처리 완료 ({status})"  # 메시지 생성
                progress_cb(percent, msg)  # 콜백 함수로 알림

    # 전체 마스터 합치기
    if per_file_dfs:  # 데이터가 있으면
        master = pd.concat(per_file_dfs, ignore_index=True)  # 여러 파일 데이터 합침
        master = master.drop_duplicates(subset=["category_id"]).reset_index(drop=True)  # 중복 제거
    else:  # 데이터가 없으면
        master = pd.DataFrame(
            columns=["category_id", "category_path", "level1", "level2", "level3", "level4"]  # 빈 테이블 생성
        )

    master.to_pickle(MASTER_CACHE_FILE)  # 전체 결과를 캐시 파일로 저장

    if progress_cb:  # 콜백 함수가 있으면
        progress_cb(100, "카테고리 마스터 생성 완료")  # 100% 완료 메시지 알림

    return master  # 최종 결과 반환

# === 3) 외부에서 쓰기 편하게: 마스터 로드 헬퍼 ===
_category_master_cache: pd.DataFrame | None = None  # 내부 캐시


def load_category_master(force_rebuild: bool = False,
                         progress_cb: Optional[Callable[[int, str], None]] = None) -> pd.DataFrame:
    """
    - force_rebuild=False 이면:
        1) 메모리 캐시 있으면 그대로 사용
        2) 메모리 캐시 없고, MASTER_CACHE_FILE 이 있으면 pkl 로드
        3) pkl 도 없으면 get_category_master() 돌려서 새로 생성
    - force_rebuild=True 이면:
        무조건 get_category_master()로 다시 만들고 pkl 갱신
    """
    global _category_master_cache

    # 1) 메모리 캐시 우선
    if _category_master_cache is not None and not force_rebuild:
        return _category_master_cache

    # 2) 파일 캐시(pkl) 사용
    if not force_rebuild and MASTER_CACHE_FILE.exists():
        if progress_cb:
            progress_cb(0, f"기존 카테고리 마스터 캐시 로드: {MASTER_CACHE_FILE}")
        df = pd.read_pickle(MASTER_CACHE_FILE)
        if progress_cb:
            progress_cb(100, f"카테고리 마스터 캐시 로드 완료 (총 {len(df)}개)")
        _category_master_cache = df
        return df

    # 3) 없으면 새로 생성
    df = get_category_master(category_dir=CATEGORY_DIR, progress_cb=progress_cb)
    _category_master_cache = df
    return df



if __name__ == "__main__":  # 이 파일을 직접 실행할 때만 아래 코드 실행
    def debug_log(p, m):  # 진행상황을 출력하는 함수
        print(f"{p}% | {m}")

    df = get_category_master(progress_cb=debug_log)  # 마스터 데이터 생성, 진행상황 출력
    print("총 카테고리 수:", len(df))  # 전체 카테고리 수 출력


