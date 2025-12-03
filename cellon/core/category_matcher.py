# cellon/core/category_matcher.py

from __future__ import annotations

from typing import Dict, Any, Optional

import pandas as pd

from cellon.category_ai.category_loader import load_category_master
from cellon.category_ai.category_llm import suggest_category_with_candidates
from cellon.core.rules_loader import load_meta_kitchen_rules, load_coupang_kitchen_rules


class CategoryMatcher:
    def __init__(self) -> None:
        self.cat_master = load_category_master()
        self.meta_rules = load_meta_kitchen_rules()
        self.coupang_rules = load_coupang_kitchen_rules()

    # --- 메타 카테고리 추론 ---
    def _infer_meta_key(
        self,
        source: str,
        source_category_path: str,
        name: str,
    ) -> Optional[str]:
        """
        소스몰/카테고리 path/상품명으로 meta_kitchen_* 키 추론
        """
        source = source.lower()
        path = (source_category_path or "").strip()

        # 1) path 기준으로 먼저 매칭
        for meta_key, rule in self.meta_rules.items():
            if source == "costco":
                paths = rule.get("source_costco_paths", [])
            elif source == "domemae":
                paths = rule.get("source_domemae_paths", [])
            else:
                paths = []

            for p in paths:
                if path == p:
                    return meta_key

        # 2) 키워드 기준 보조 매칭
        name_lower = name.lower()
        for meta_key, rule in self.meta_rules.items():
            inc = rule.get("keywords_include", [])
            exc = rule.get("keywords_exclude", [])
            if not inc:
                continue

            if any(k in name for k in inc) and not any(k in name for k in exc):
                return meta_key

        return None

    # --- 외부에서 쓰는 진입점 ---
    def match_category(
        self,
        source: str,
        source_category_path: str,
        product_name: str,
        brand: Optional[str] = None,
        extra_text: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        반환 예:
        {
          "category_id": "12345",
          "category_path": "주방용품>냄비/솥",
          "reason": "..."
        }
        """
        meta_key = self._infer_meta_key(source, source_category_path, product_name)

        if meta_key is None:
            # 룰로 못 골랐으면 전체 cat_master를 후보 없이 LLM에 넘김
            return suggest_category_with_candidates(
                product_name=product_name,
                brand=brand,
                extra_text=extra_text,
                candidates_df=None,
            )

        coupang_rule = self.coupang_rules.get(meta_key, {})
        candidate_ids = coupang_rule.get("coupang_category_ids", [])

        if not candidate_ids:
            return suggest_category_with_candidates(
                product_name=product_name,
                brand=brand,
                extra_text=extra_text,
                candidates_df=None,
            )

        candidates_df = self.cat_master[
            self.cat_master["category_id"].astype(str).isin(candidate_ids)
        ]

        if len(candidates_df) == 1:
            row = candidates_df.iloc[0]
            return {
                "category_id": row["category_id"],
                "category_path": row["category_path"],
                "reason": f"meta_key={meta_key} 룰에 의해 단일 후보 자동 선택",
            }

        # 2개 이상이면 후보 제한 모드 LLM 호출
        return suggest_category_with_candidates(
            product_name=product_name,
            brand=brand,
            extra_text=extra_text,
            candidates_df=candidates_df,
        )
