# src/build_coupang_upload_index.py

from pathlib import Path
import json

from cellon.config import COUPANG_UPLOAD_FORM_DIR, COUPANG_UPLOAD_INDEX_JSON


def build_coupang_upload_index() -> Path:
    """
    coupang_upload_form 폴더를 한 번 훑어서
    sellertool_upload_*.xlsm 템플릿에 대한 인덱스 JSON을 생성한다.

    이 스크립트는 수동으로(또는 나중에 '카테고리 분석 시작' 버튼에서)
    한 번만 실행해주면 된다.
    """
    root = COUPANG_UPLOAD_FORM_DIR

    if not root.exists():
        raise FileNotFoundError(f"쿠팡 업로드 폼 루트 폴더가 없습니다: {root}")

    templates = []

    # 폴더 전체를 재귀적으로 탐색
    for path in root.rglob("sellertool_upload_*.xlsm"):
        key = path.stem.replace("sellertool_upload_", "")
        rel_path = path.relative_to(root)

        templates.append({
            "key": key,
            "relative_path": str(rel_path),
        })

    if not templates:
        raise RuntimeError(f"쿠팡 업로드 폼 템플릿을 찾지 못했습니다: {root}")

    # JSON 저장
    COUPANG_UPLOAD_INDEX_JSON.parent.mkdir(parents=True, exist_ok=True)
    data = {"templates": templates}

    with COUPANG_UPLOAD_INDEX_JSON.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"[OK] 쿠팡 템플릿 인덱스 {len(templates)}개 생성")
    print(f"     JSON 경로: {COUPANG_UPLOAD_INDEX_JSON}")

    return COUPANG_UPLOAD_INDEX_JSON


if __name__ == "__main__":
    build_coupang_upload_index()
