# cellon/category_ai/category_rules_builder.py

from __future__ import annotations
from pathlib import Path
import json
import pandas as pd

from .category_loader import load_category_master

# ① 여기만 추가/수정하면 됨
#    - group 별로 meta_key → 매칭할 category_path 부분 문자열 리스트
GROUP_RULE_CONFIG: dict[str, dict[str, list[str]]] = {
    "kitchen": {
        # meta_kitchen_pot: '주방용품>냄비/솥' 들어가는 카테고리 다 모으기
        "meta_kitchen_pot": [
            "주방용품>냄비/솥",
        ],
        # meta_kitchen_pan: '주방용품>프라이팬/그릴' 들어가는 카테고리 다 모으기
        "meta_kitchen_pan": [
            "주방용품>프라이팬/그릴",
        ],
    },

    # 예시: 나중에 food 그룹도 추가 가능
    # "food": {
    #     "meta_food_snack": ["가공식품>과자/스낵", "가공식품>쿠키"],
    #     "meta_food_instant": ["가공식품>라면", "가공식품>즉석식품"],
    # },
    #
    # "beauty": { ... },
}


def _collect_ids_for_paths(df: pd.DataFrame, path_substrings: list[str]) -> list[str]:
    """category_path 에 특정 문자열이 포함된 category_id 목록 수집"""
    if not path_substrings:
        return []

    mask = False
    col = df["category_path"].astype(str)

    for sub in path_substrings:
        mask = mask | col.str.contains(sub, na=False)

    ids = (
        df.loc[mask, "category_id"]
        .astype(str)
        .dropna()
        .unique()
        .tolist()
    )
    return ids


def build_coupang_rules_for_all_groups(df: pd.DataFrame | None = None) -> None:
    """
    - category_master.pkl 을 읽거나(df가 주어지면 그대로 사용)
    - GROUP_RULE_CONFIG 에 정의된 모든 group 에 대해
      cellon/rules/coupang/{group}_rules.json 을 자동 생성/갱신.
    - 기존 파일에 있던 "__strong_name_rules__" 는 그대로 보존한다.
    """
    # 1) df 없으면 pkl 로드
    if df is None:
        df = load_category_master()

    # 2) 프로젝트 루트 및 coupang rules 디렉토리
    project_root = Path(__file__).resolve().parents[2]  # .../Project_cellon (현재 구조 기준)
    coupang_rules_dir = project_root / "cellon" / "rules" / "coupang"
    coupang_rules_dir.mkdir(parents=True, exist_ok=True)

    for group, meta_config in GROUP_RULE_CONFIG.items():
        new_rules: dict[str, dict] = {}

        # 1) meta_* 룰 새로 계산
        for meta_key, path_substrings in meta_config.items():
            ids = _collect_ids_for_paths(df, path_substrings)
            new_rules[meta_key] = {
                "coupang_category_ids": ids,
                "priority": 100,
            }
            print(
                f"[rules_builder] group={group}, meta_key={meta_key}, "
                f"매칭된 카테고리 수={len(ids)}"
            )

        out_path = coupang_rules_dir / f"{group}_rules.json"  # kitchen → kitchen_rules.json

        # 2) 기존 strong_name_rules 보존
        existing: dict = {}
        if out_path.exists():
            try:
                with out_path.open("r", encoding="utf-8") as f:
                    existing = json.load(f) or {}
                if not isinstance(existing, dict):
                    existing = {}
            except Exception:
                existing = {}

        STRONG_KEY = "__strong_name_rules__"
        if STRONG_KEY in existing:
            # 기존 strong_name_rules 그대로 붙여넣기
            new_rules[STRONG_KEY] = existing[STRONG_KEY]

        # 3) 최종 저장
        with out_path.open("w", encoding="utf-8") as f:
            json.dump(new_rules, f, ensure_ascii=False, indent=2)

        print(
            f"✅ {group}_rules.json 생성/갱신 완료 "
            f"(strong_name_rules 보존) → {out_path}"
        )


def main():
    df = load_category_master()
    print(f"[rules_builder] category_master 로드 완료: {len(df)}개 카테고리")
    build_coupang_rules_for_all_groups(df)


if __name__ == "__main__":
    main()
