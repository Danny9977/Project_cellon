# src/build_coupang_upload_index.py

from pathlib import Path
import json

from cellon.config import COUPANG_UPLOAD_FORM_DIR, COUPANG_UPLOAD_INDEX_JSON


# src/build_coupang_upload_index.py

from pathlib import Path
import json

from cellon.config import COUPANG_UPLOAD_FORM_DIR, COUPANG_UPLOAD_INDEX_JSON


def build_coupang_upload_index() -> Path:
    """
    coupang_upload_form 폴더를 한 번 훑어서
    sellertool_upload_*.xlsm 템플릿에 대한 인덱스 JSON을 생성한다.

    - 템플릿 파일은 rglob()로 재귀 탐색한다 (A 방식)
    - 생성된 인덱스 JSON을 템플릿 선택 단계에서 우선 사용한다 (B 방식)
    """
    root = COUPANG_UPLOAD_FORM_DIR

    if not root.exists():
        raise FileNotFoundError(f"쿠팡 업로드 폼 루트 폴더가 없습니다: {root}")

    templates = []
    key_to_paths: dict[str, list[str]] = {}

    for path in root.rglob("sellertool_upload_*.xlsm"):
        key = path.stem.replace("sellertool_upload_", "")
        rel_path = str(path.relative_to(root))

        templates.append({
            "key": key,
            "relative_path": rel_path,
        })

        key_to_paths.setdefault(key, []).append(rel_path)

    if not templates:
        raise RuntimeError(f"쿠팡 업로드 폼 템플릿을 찾지 못했습니다: {root}")

    # (선택) key 중복 경고: 같은 key가 여러 파일에 있으면 선택 시 혼동 가능
    duplicates = {k: v for k, v in key_to_paths.items() if len(v) > 1}
    if duplicates:
        print("[WARN] 동일 key 템플릿이 여러 개 발견되었습니다. (선택 로직에서 첫 번째로 매칭될 수 있음)")
        for k, v in list(duplicates.items())[:20]:
            print(f"  - key='{k}' -> {v}")

    COUPANG_UPLOAD_INDEX_JSON.parent.mkdir(parents=True, exist_ok=True)

    data = {
        "root": str(root),  # 참고용
        "templates": templates,
    }

    with COUPANG_UPLOAD_INDEX_JSON.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"[OK] 쿠팡 템플릿 인덱스 {len(templates)}개 생성")
    print(f"     JSON 경로: {COUPANG_UPLOAD_INDEX_JSON}")

    return COUPANG_UPLOAD_INDEX_JSON


if __name__ == "__main__":
    build_coupang_upload_index()
