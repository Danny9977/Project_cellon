# cellon/core/category_matcher.py

from __future__ import annotations

from typing import Any, Dict, Optional, Callable

import pandas as pd
import difflib 

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
        manual_resolver: Optional[
            Callable[[str, str, pd.DataFrame], Optional[Dict[str, Any]]]
        ] = None,
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
        
        # 수동 해결기 (디버깅/테스트용)
        self._manual_resolver = manual_resolver

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

        # --- 1차 룰 매칭 실패: meta_key 자체가 없어도 strong_name_rule 물어보기 ---
        if meta_key is None:
            self._log("  ▶ meta_key=None → 1차 룰 매칭 실패")

            # 0) strong_name_rules 대상 카테고리만 후보로 구성
            strong_rules = self.coupang_rules.get("__strong_name_rules__", []) or []
            strong_ids = {
                str(rule.get("target_category_id"))
                for rule in strong_rules
                if rule.get("target_category_id")
            }

            strong_candidates_df = self.cat_master[
                self.cat_master["category_id"].astype(str).isin(strong_ids)
            ]
            self._log(f"  ▶ strong_name_rules 기반 후보 수: {len(strong_candidates_df)}")

            # 0-1) strong_name_rules 매칭 시도
            strong_result = self._pick_by_strong_keyword(
                product_name=product_name,
                candidates_df=strong_candidates_df,
            )
            if strong_result is not None:
                strong_result.update(
                    {
                        "used_llm": False,
                        "meta_key": None,
                        "num_candidates": len(strong_candidates_df),
                    }
                )
                self._log(
                    f"  🔚 strong_name_rules 결과 사용 "
                    f"(meta_key=None, LLM 미호출): {strong_result}"
                )
                return strong_result

            # ✅ 0-2) strong 후보들에 대해 수동 선택 기회 제공
            if (
                self._manual_resolver is not None
                and strong_candidates_df is not None
                and not strong_candidates_df.empty
            ):
                self._log("  ▶ strong 후보들에 대해 수동 카테고리 선택 요청 (meta_key=None)")
                manual = self._manual_resolver(
                    product_name,
                    source_category_path,
                    strong_candidates_df,
                )
                if manual is not None:
                    manual.setdefault("used_llm", False)
                    manual.setdefault("meta_key", None)
                    manual.setdefault("num_candidates", len(strong_candidates_df))
                    self._log(
                        f"  🔚 수동 선택 결과 사용 "
                        f"(meta_key=None, LLM 미호출): {manual}"
                    )
                    return manual
                self._log(
                    "  ▶ 수동 선택 없음 또는 "
                    "'LLM에게 맡기기' 선택 → 전체 LLM 진행"
                )

            # (원래 있던 전체 LLM 호출은 이 아래로 내려감)
            self._log("  ▶ strong_name_rules/수동 선택 실패 → 전체 cat_master LLM 검색 모드")
            self._log(f"    - 전체 카테고리 수: {len(self.cat_master)}")

            llm_result = suggest_category_with_candidates(
                product_name=product_name,
                brand=brand,
                extra_text=extra_text,
                candidates_df=None,
            )
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

        # 쿠팡 카테고리 마스터에서 후보 필터링
        candidates_df = self.cat_master[
            self.cat_master["category_id"].astype(str).isin([str(cid) for cid in candidate_ids])
        ]
        self._log(f"  ▶ 후보 cat_master 필터링 완료: {len(candidates_df)}개 행")

        # ✅ 0단계: strong_name_rules 먼저 시도
        strong_result = self._pick_by_strong_keyword(
            product_name=product_name,
            candidates_df=candidates_df,
        )
        if strong_result is not None:
            strong_result.update(
                {
                    "used_llm": False,
                    "meta_key": meta_key,
                    "num_candidates": len(candidates_df),
                }
            )
            self._log(f"  🔚 strong_name_rules 결과 사용 (LLM 미호출): {strong_result}")
            return strong_result


        # 1차: leaf 완전 일치 룰
        leaf_result = self._pick_by_leaf_keyword(
            product_name=product_name,
            candidates_df=candidates_df,
        )
        if leaf_result is not None:
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

        # 2차: 후보가 1개면 그냥 확정
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

        # 🔹 3차: 후보가 2개 이상 → 먼저 사람에게 물어보기
        if self._manual_resolver is not None and not candidates_df.empty:
            self._log("  ▶ 수동 카테고리 선택 콜백 호출 (LLM 이전 단계)")
            manual = self._manual_resolver(
                product_name,
                source_category_path,
                candidates_df,
            )
            if manual is not None:
                # manual 예: {"category_id":"80298", "category_path":"주방용품>...>웍", "reason":"사용자 수동 선택"}
                manual.setdefault("used_llm", False)
                manual.setdefault("meta_key", meta_key)
                manual.setdefault("num_candidates", len(candidates_df))
                self._log(f"  🔚 수동 선택 결과 사용 (LLM 미호출): {manual}")
                return manual
            self._log("  ▶ 수동 선택 없음 또는 'LLM에게 맡기기' 선택 → LLM 진행")

        # 🔹 4차: 그래도 결정 안되면 LLM에게
        self._log(f"  ▶ 후보 {len(candidates_df)}개 → 제한된 후보만 LLM에 전달")
        llm_raw = suggest_category_with_candidates(
            product_name=product_name,
            brand=brand,
            extra_text=source_category_path,
            candidates_df=candidates_df,
        )
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
        # ======= 내부 헬퍼: leaf 정규화 및 매칭 =======
    @staticmethod
    def _normalize_leaf_text(text: str) -> str:
        """
        카테고리 leaf 이름 / 키워드를 비교하기 위한 정규화:
        - 공백/기호 제거
        - 한글/영문/숫자만 남기기
        """
        text = text or ""
        text = re.sub(r"[^0-9a-zA-Z가-힣]", "", text)
        return text.lower()

    @staticmethod
    def _char_similarity(a: str, b: str) -> float:
        """
        문자 단위 유사도 (0.0 ~ 1.0)
        - 1.0: 완전 동일
        - 0.8 이상: 상당히 비슷
        """
        if not a or not b:
            return 0.0
        return difflib.SequenceMatcher(None, a, b).ratio()

    def _pick_by_leaf_keyword(
        self,
        product_name: str,
        candidates_df: pd.DataFrame,
    ) -> Optional[Dict[str, Any]]:
        """
        1차 룰 (강한 확신이 있을 때만 확정):

        - 상품명에서 키워드 추출 (_extract_keywords 사용)
        - 정규화 + 연속 단어 bigram 을 만들어 토큰 후보 생성
        - 각 후보 category_path 의 마지막 뎁스(leaf)와 문자 유사도 계산
        - best_score 가 threshold 이상이고, 2등과의 차이가 margin 이상이면
          → 1차에서 확정
        - 나머지는 LLM에게 위임 (None 반환)
        """
        if candidates_df is None or candidates_df.empty:
            return None

        # 1) 상품명 키워드 추출
        raw_keywords = _extract_keywords(product_name, brand=None, extra=None)
        if not raw_keywords:
            return None

        # 정규화된 토큰
        norm_tokens = [
            self._normalize_leaf_text(k)
            for k in raw_keywords
        ]
        norm_tokens = [t for t in norm_tokens if t]

        if not norm_tokens:
            return None

        # 2) 연속 단어 bigram 합성 (띄어쓰기 보정: 편수 + 냄비 → 편수냄비)
        bigram_tokens: list[str] = []
        for i in range(len(norm_tokens) - 1):
            bigram_tokens.append(norm_tokens[i] + norm_tokens[i + 1])

        all_tokens = norm_tokens + bigram_tokens
        if not all_tokens:
            return None

        # 3) 각 후보 leaf와의 문자 유사도 계산
        best = None
        best_score = 0.0
        second_best = 0.0

        for _, row in candidates_df.iterrows():
            cat_id = str(row["category_id"])
            cat_path = str(row["category_path"])

            # 마지막 뎁스 추출 (예: '주방용품>취사도구>냄비>편수냄비' → '편수냄비')
            leaf = cat_path.split(">")[-1].strip()
            leaf_norm = self._normalize_leaf_text(leaf)

            if not leaf_norm:
                continue

            # 이 leaf에 대해 가장 잘 맞는 토큰의 유사도
            leaf_best = 0.0
            for token in all_tokens:
                score = self._char_similarity(leaf_norm, token)
                if score > leaf_best:
                    leaf_best = score

            if leaf_best > best_score:
                second_best = best_score
                best_score = leaf_best
                best = {
                    "category_id": cat_id,
                    "category_path": cat_path,
                    "leaf": leaf,
                }
            elif leaf_best > second_best:
                second_best = leaf_best

        # 4) 최종 결정 조건
        # - threshold 이상: 꽤 확신 있는 경우만
        # - margin 이상: 1등과 2등 차이가 일정 수준 이상
        THRESHOLD = 0.75   # 너무 낮추지 않는 게 좋습니다 (0.8~0.9 사이 추천)
        MARGIN = 0.05      # 5% 정도 차이는 나야 "단일 승자"로 인정

        if best and best_score >= THRESHOLD and (best_score - second_best) >= MARGIN:
            self._log(
                "  ✅ [1차 룰] leaf-문자 유사도로 카테고리 확정 "
                f"(best={best_score:.2f}, second={second_best:.2f})"
            )
            self._log(
                f"     - leaf='{best['leaf']}', "
                f"category_id={best['category_id']}, path={best['category_path']}"
            )
            return {
                "category_id": best["category_id"],
                "category_path": best["category_path"],
                "reason": (
                    "상품명 키워드(단어/연속 단어)의 문자 유사도가 "
                    "쿠팡 카테고리 마지막 뎁스와 매우 높습니다. "
                    f"(similarity={best_score:.2f}, leaf='{best['leaf']}')"
                ),
                "used_llm": False,
            }

        if best:
            self._log(
                "  ℹ️ [1차 룰] leaf-문자 유사도 후보는 있었지만 "
                "threshold 미달 또는 2등과의 차이가 작아서 LLM으로 위임 "
                f"(best={best_score:.2f}, second={second_best:.2f})"
            )

        return None
    
    # --------- ----- 내부 헬퍼: 강제 name 룰 매칭 ---------
    def _pick_by_strong_keyword(
        self,
        product_name: str,
        candidates_df: pd.DataFrame,
    ) -> Optional[Dict[str, Any]]:
        """
        강제 name 룰:
        - coupang/kitchen_rules.json 의 "__strong_name_rules__" 를 읽어서
        - 상품명에 특정 키워드가 포함되고
        - 그 target_category_id 가 현재 후보(candidates_df)에 포함되어 있으면
          → 그 카테고리로 바로 확정.
        """
        if candidates_df is None or candidates_df.empty:
            return None

        rules = self.coupang_rules.get("__strong_name_rules__", []) or []
        if not rules:
            return None

        name_lower = (product_name or "").lower()
        # 후보 ID 셋 (문자열 기준)
        candidate_ids = set(candidates_df["category_id"].astype(str).tolist())

        matches: list[Dict[str, Any]] = []

        for rule in rules:
            keywords = rule.get("keywords") or []
            target_id = str(rule.get("target_category_id") or "").strip()
            if not target_id:
                continue

            # 현재 meta_key에서 뽑힌 후보들 안에 target_id가 없으면 스킵
            if target_id not in candidate_ids:
                continue

            # 키워드 중 하나라도 상품명에 포함되면 매칭
            if any(kw.lower() in name_lower for kw in keywords):
                matches.append(
                    {
                        "target_id": target_id,
                        "rule": rule,
                    }
                )

        if len(matches) == 1:
            m = matches[0]
            target_id = m["target_id"]
            rule = m["rule"]

            row = candidates_df[candidates_df["category_id"].astype(str) == target_id].iloc[0]
            cat_path = str(row["category_path"])

            self._log(
                "  ✅ [강제 name 룰] 키워드에 의해 카테고리 확정: "
                f"target_id={target_id}, path={cat_path}, rule_keywords={rule.get('keywords')}"
            )

            return {
                "category_id": target_id,
                "category_path": cat_path,
                "reason": rule.get("reason")
                or f"강제 name 룰에 의해 '{rule.get('keywords')}' 키워드로 카테고리를 결정했습니다.",
                "used_llm": False,
            }

        if len(matches) > 1:
            # 여러 룰이 동시에 걸리면 모호하니 LLM/leaf 룰로 넘김
            self._log("  ⚠️ [강제 name 룰] 매칭 룰이 여러 개 → LLM/leaf 룰로 위임")
            for m in matches:
                self._log(
                    f"     - target_id={m['target_id']}, rule_keywords={m['rule'].get('keywords')}"
                )

        return None

