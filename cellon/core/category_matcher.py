# cellon/core/category_matcher.py

from __future__ import annotations

from typing import Any, Dict, Optional, Callable

import pandas as pd

from cellon.category_ai.category_loader import load_category_master
from cellon.core.rules_loader import (
    load_meta_rules,
    load_coupang_rules,
    list_available_groups,
    load_all_market_groups,
)

import re

from cellon.category_ai.category_llm import (
    suggest_category_with_candidates,
    _extract_keywords,
)


class CategoryMatcher:
    """
    - group(예: "kitchen", "food", "beauty") 단위로 meta / coupang 룰 자동 로딩
    - rules/<market>/<group>.json 을 자동 스캔해서 source(코코/도매매/오너) 카테고리 정보도 함께 들고 있음
    - 1차: meta 룰 기반 필터링
    - 실패 시: LLM에게 전체/부분 후보 전달
    """

    def __init__(
        self,
        group: str = "kitchen",
        logger: Optional[Callable[[str], None]] = None,
    ) -> None:
        """
        Parameters
        ----------
        group:
            사용할 카테고리군 이름.
            - "kitchen" (기본값)
            - "food", "beauty" 등도 meta / coupang 룰 파일만 추가하면 그대로 동작
        """
        self.group = group

        # UI에서 넘겨주는 logger (예: ChromeCrawler._log)
        self._logger = logger

        # 1) 쿠팡 전체 카테고리 마스터 (엑셀 → pkl)
        self.cat_master: pd.DataFrame = load_category_master()

        # 2) meta / coupang 룰 (group 기반 자동 로딩)
        self.meta_rules: Dict[str, Any] = load_meta_rules(group)
        self.coupang_rules: Dict[str, Any] = load_coupang_rules(group)

        # 3) rules/<market>/<group>.json 전체 (디버깅/추가 로직용)
        #    예: self.market_trees["costco"]["kitchen"]["categories"] ...
        self.market_trees: Dict[str, Dict[str, Dict[str, Any]]] = load_all_market_groups()

        self._log(f"[CategoryMatcher] 초기화 완료: group={group}")
        self._log(f"  - meta_rules: {len(self.meta_rules)}개")
        self._log(f"  - coupang_rules: {len(self.coupang_rules)}개")
        self._log(f"  - market_trees: {list(self.market_trees.keys())}")
        
    # 내부용 로그 헬퍼
    def _log(self, msg: str) -> None:
        try:
            if self._logger:
                self._logger(msg)
        except Exception:
            # logger 쪽 문제로 매칭이 죽지 않도록 방어
            pass
    
    # ------------------------------------------------------------------
    # 헬퍼: 사용 가능한 group 목록 확인
    # ------------------------------------------------------------------
    @staticmethod
    def available_groups() -> list[str]:
        """현재 meta/coupang 룰 기준으로 사용 가능한 group 목록."""
        return list_available_groups()

    # ------------------------------------------------------------------
    # 1) 메타 카테고리 추론
    # ------------------------------------------------------------------
    def _infer_meta_key(
        self,
        source: str,
        source_category_path: str,
        name: str,
    ) -> Optional[str]:
        """
        소스몰(source)/카테고리 path/상품명으로 meta_kitchen_* 같은 meta key 추론.

        - 1순위: path 완전 일치 (coupang_*.json 내 source_*_paths)
        - 2순위: keywords_include / keywords_exclude 기반 부분 매칭
        """
        source = (source or "").lower()
        path = (source_category_path or "").strip()
        name = name or ""
        name_lower = name.lower()

        self._log("────────────────────────────────────────")
        self._log("[CategoryMatcher] _infer_meta_key 호출")
        self._log(f"  - source={source}")
        self._log(f"  - path='{path}'")
        self._log(f"  - name='{name}'")

        # 1) path 기준으로 먼저 매칭
        self._log("  ▶ 1단계: path 완전 일치 매칭 시도")
        
        for meta_key, rule in self.meta_rules.items():
            if source == "costco":
                paths = rule.get("source_costco_paths", [])
            elif source == "domemae":
                paths = rule.get("source_domemae_paths", [])
            elif source == "owner":
                paths = rule.get("source_owner_paths", [])
            else:
                paths = []

            if paths:
                self._log(f"    - meta_key={meta_key} / paths={paths}")

            for p in paths:
                if path == p:
                    self._log(f"  ✅ path 매칭 성공: meta_key={meta_key} (path='{p}')")
                    return meta_key

        # 2) 키워드 기준 보조 매칭
        self._log("  ▶ 2단계: keywords_include / exclude 매칭 시도")
        for meta_key, rule in self.meta_rules.items():
            inc = rule.get("keywords_include", []) or []
            exc = rule.get("keywords_exclude", []) or []

            if not inc:
                continue

            self._log(f"    - meta_key={meta_key} / include={inc} / exclude={exc}")
            
            if any(k in name_lower for k in inc) and not any(k in name_lower for k in exc):
                self._log(f"  ✅ 키워드 매칭 성공: meta_key={meta_key}")
                return meta_key

        self._log("  ❌ meta_key 추론 실패 (path/키워드 모두 불일치)")
        return None

    # ------------------------------------------------------------------
    # 2) 외부에서 사용하는 진입점
    # ------------------------------------------------------------------
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
            "reason": "...",
            "used_llm": False,
            "meta_key": "meta_kitchen_pot",
            "num_candidates": 1,
        }

        흐름:
        1) meta 룰로 meta_key 추론
        2) meta_key -> coupang_rules[group] 에서 candidate_ids 추출
        3) 후보가 0개면 → 전체 cat_master 를 LLM 에 그대로 넘김
        4) 후보가 1개면 → 바로 선택
        5) 후보가 여러 개면 → 제한된 candidates_df 만 LLM 에 넘김
        """
        
        self._log("════════════════════════════════════════")
        self._log("[CategoryMatcher] match_category 시작")
        self._log(f"  - group={self.group}")
        self._log(f"  - source={source}")
        self._log(f"  - source_category_path='{source_category_path}'")
        self._log(f"  - product_name='{product_name}'")
        
        meta_key = self._infer_meta_key(source, source_category_path, product_name)

        # --- 1차 룰 매칭 실패: meta_key 자체가 없으면 → 전체 LLM 검색 ---
        if meta_key is None:
            self._log("  ▶ meta_key=None → 1차 룰 매칭 실패 → 전체 cat_master LLM 검색 모드")
            self._log(f"    - 전체 카테고리 수: {len(self.cat_master)}")

            llm_result = suggest_category_with_candidates(
                product_name=product_name,
                brand=brand,
                extra_text=extra_text,
                candidates_df=None,
            )
            # LLM 호출임을 표기
            llm_result.setdefault("used_llm", True)
            llm_result.setdefault("meta_key", None)
            llm_result.setdefault("num_candidates", None)

            self._log(f"  🔚 LLM 결과 수신 (meta_key=None): {llm_result}")
            return llm_result

        self._log(f"  ▶ meta_key 추론 성공: {meta_key}")
        coupang_rule = self.coupang_rules.get(meta_key, {})
        candidate_ids = coupang_rule.get("coupang_category_ids", []) or []
        self._log(f"    - coupang_rule.coupang_category_ids = {candidate_ids}")

        # meta_key는 있는데, 거기에 매핑된 후보가 없으면 → 전체 LLM 검색
        if not candidate_ids:
            self._log("  ❌ meta_key는 있으나 candidate_ids가 비어 있음 → 전체 cat_master LLM 검색")
            self._log(f"    - 전체 카테고리 수: {len(self.cat_master)}")
            
            llm_result = suggest_category_with_candidates(
                product_name=product_name,
                brand=brand,
                extra_text=extra_text,
                candidates_df=None,
            )
            llm_result.setdefault("used_llm", True)
            llm_result.setdefault("meta_key", meta_key)
            llm_result.setdefault("num_candidates", None)

            self._log(f"  🔚 LLM 결과 수신 (candidate_ids 없음): {llm_result}")
            return llm_result

        # 쿠팡 카테고리 마스터에서 해당 ID 들만 필터링
        candidates_df = self.cat_master[
            self.cat_master["category_id"].astype(str).isin([str(cid) for cid in candidate_ids])
        ]
        self._log(f"  ▶ 후보 cat_master 필터링 완료: {len(candidates_df)}개 행")

        # ✅ 1차: leaf(마지막 뎁스) 완전 일치 룰 시도
        leaf_result = self._pick_by_leaf_keyword(
            product_name=product_name,
            candidates_df=candidates_df,
        )

        if leaf_result is not None:
            # 1차 룰로 이미 확정된 경우 → LLM은 부르지 않음
            result = {
                "category_id": leaf_result["category_id"],
                "category_path": leaf_result["category_path"],
                "reason": leaf_result["reason"],
                "used_llm": False,
                "meta_key": meta_key,
                "num_candidates": len(candidates_df),
            }
            self._log(f"  🔚 1차 leaf 룰 결과 사용 (LLM 미호출): {result}")
            return result

        # ✅ 2차: 후보가 1개뿐이면 → 룰만으로 결정 (LLM 미사용)
        if len(candidates_df) == 1:
            row = candidates_df.iloc[0]
            result = {
                "category_id": str(row["category_id"]),
                "category_path": str(row["category_path"]),
                "reason": f"group={self.group}, meta_key={meta_key} 룰에 의해 단일 후보 자동 선택",
                "used_llm": False,
                "meta_key": meta_key,
                "num_candidates": 1,
            }
            self._log(f"  🔚 단일 후보 자동 선택 (LLM 미호출): {result}")
            return result

        # ✅ 3차: 그 외(후보 2개 이상이고 leaf 룰도 실패) → 제한된 후보만 LLM에 전달
        self._log(f"  ▶ 후보 {len(candidates_df)}개 → 제한된 후보만 LLM에 전달")

        llm_raw = suggest_category_with_candidates(
            product_name=product_name,
            brand=brand,                     # 필요하면 brand 그대로 넘김
            extra_text=source_category_path, # ✅ 여기가 source_path가 아니라 source_category_path
            candidates_df=candidates_df,
        )

        # CategoryMatcher 스타일에 맞게 공통 필드 보정
        llm_raw.update(
            {
                "used_llm": True,
                "meta_key": meta_key,
                "num_candidates": len(candidates_df),
            }
        )

        self._log(f"  🔚 LLM 결과 수신 (제한 후보): {llm_raw}")
        return llm_raw


    
    # ======= 내부 헬퍼: leaf 정규화 및 매칭 =======
    @staticmethod
    def _normalize_leaf_text(text: str) -> str:
        """
        카테고리 leaf 이름 / 키워드를 비교하기 위한 정규화:
        - 공백 제거
        - 한글/영문/숫자만 남기기
        """
        text = text or ""
        # 한글/영문/숫자만 남기고 나머지는 제거
        text = re.sub(r"[^0-9a-zA-Z가-힣]", "", text)
        return text.lower()

    def _pick_by_leaf_keyword(
        self,
        product_name: str,
        candidates_df: pd.DataFrame,
    ) -> Optional[Dict[str, Any]]:
        """
        1차 룰:
        - 상품명에서 키워드 추출
        - 각 후보 category_path 의 마지막 뎁스(leaf)를 비교
        - leaf 와 키워드가 '완전히 일치'하는 후보가 딱 1개이면, 그걸 바로 선택
        - 여러 개면 애매하므로 LLM에 넘기도록 None 반환
        """
        if candidates_df is None or candidates_df.empty:
            return None

        # 상품명에서 키워드 추출 (예: '쿡에버 올인원 편수 냄비 3P' → ['쿡에버','올인원','편수','냄비','채망포함'])
        keywords = _extract_keywords(product_name, brand=None, extra=None)

        # 키워드를 정규화 (공백/기호 제거)
        kw_norms = {
            self._normalize_leaf_text(kw)
            for kw in keywords
        }
        if not kw_norms:
            return None

        matches = []

        for _, row in candidates_df.iterrows():
            cat_id = str(row["category_id"])
            cat_path = str(row["category_path"])

            # 카테고리 마지막 뎁스만 추출 (예: '주방용품>취사도구>냄비>편수냄비' → '편수냄비')
            leaf = cat_path.split(">")[-1].strip()
            leaf_norm = self._normalize_leaf_text(leaf)

            if leaf_norm and leaf_norm in kw_norms:
                matches.append(
                    {
                        "category_id": cat_id,
                        "category_path": cat_path,
                        "leaf": leaf,
                    }
                )

        if len(matches) == 1:
            m = matches[0]
            self._log("  ✅ [1차 룰] leaf 완전 일치로 카테고리 확정")
            self._log(f"     - leaf='{m['leaf']}', category_id={m['category_id']}, path={m['category_path']}")

            return {
                "category_id": m["category_id"],
                "category_path": m["category_path"],
                "reason": f"상품명 키워드와 쿠팡 카테고리 마지막 뎁스가 완전히 일치합니다. (leaf='{m['leaf']}')",
                "used_llm": False,
            }

        if len(matches) > 1:
            # 여러 후보가 leaf 기준으로 겹치면 애매하니 LLM으로 넘김
            self._log("  ⚠️ [1차 룰] leaf 일치 후보가 여러 개 → LLM으로 위임")
            for m in matches:
                self._log(f"     - 후보 leaf='{m['leaf']}', category_id={m['category_id']}, path={m['category_path']}")

        return None