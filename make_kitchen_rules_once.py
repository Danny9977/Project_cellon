# make_kitchen_rules_once.py
from pathlib import Path
import json

import pandas as pd


# === 경로 설정 ===
BASE_DIR = Path(__file__).resolve().parent

PKL_PATH = BASE_DIR / "cache" / "category_master.pkl"
RULES_PATH = BASE_DIR / "cellon" / "rules" / "coupang" / "kitchen_rules.json"


def main():
    print("🔧 [kitchen_rules] 규칙 자동 생성 시작")
    print(f"  - PKL 경로 : {PKL_PATH}")
    print(f"  - JSON 경로: {RULES_PATH}")

    if not PKL_PATH.exists():
        print("❌ category_master.pkl 이 존재하지 않습니다.")
        return

    # 1) 마스터 로드
    df = pd.read_pickle(PKL_PATH)
    print(f"✅ category_master 로드 완료 (행 수: {len(df)})")
    print(f"  - 컬럼: {df.columns.tolist()}")

    if "category_path" not in df.columns:
        print("❌ DataFrame 에 'category_path' 컬럼이 없습니다.")
        return

    # ⚠️ 여기서 ID 컬럼 이름 확인 필요: 보통 'category_id' 라고 가정
    id_col_candidates = ["category_id", "categoryId", "id"]
    id_col = None
    for c in id_col_candidates:
        if c in df.columns:
            id_col = c
            break

    if id_col is None:
        print("❌ 카테고리 ID 컬럼(category_id / categoryId / id 등)을 찾지 못했습니다.")
        return

    print(f"✅ 카테고리 ID 컬럼: '{id_col}' 사용")

    # 2) 주방용품 > 취사도구 라인만 우선 필터
    mask_kitchen = df["category_path"].str.contains("주방용품>취사도구", na=False)
    df_kitchen = df[mask_kitchen].copy()
    print(f"🔍 '주방용품>취사도구' 포함 행 수: {len(df_kitchen)}")

    # 3) 냄비 / 프라이팬 각각 별도 필터
    mask_pot = df_kitchen["category_path"].str.contains("취사도구>냄비", na=False)
    mask_pan = df_kitchen["category_path"].str.contains("취사도구>프라이팬", na=False)

    pot_df = df_kitchen[mask_pot]
    pan_df = df_kitchen[mask_pan]

    pot_ids = sorted(pot_df[id_col].astype(str).unique().tolist())
    pan_ids = sorted(pan_df[id_col].astype(str).unique().tolist())

    print(f"  ▶ meta_kitchen_pot 후보 카테고리 수: {len(pot_ids)}")
    print(f"  ▶ meta_kitchen_pan 후보 카테고리 수: {len(pan_ids)}")

    # 디버깅용: 상위 몇 개만 미리 보여주기
    print("  - POT 예시 5개:", pot_df["category_path"].head(5).to_list())
    print("  - PAN 예시 5개:", pan_df["category_path"].head(5).to_list())

    # 4) 룰 JSON 구성
    rules = {
        "meta_kitchen_pot": {
            "coupang_category_ids": pot_ids,
            "priority": 100,
        },
        "meta_kitchen_pan": {
            "coupang_category_ids": pan_ids,
            "priority": 100,
        },
    }

    RULES_PATH.parent.mkdir(parents=True, exist_ok=True)
    with RULES_PATH.open("w", encoding="utf-8") as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)

    print("💾 kitchen_rules.json 저장 완료!")
    print("✅ 규칙 생성 종료")


if __name__ == "__main__":
    main()
