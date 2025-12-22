# cellon/category_ai/category_llm.py

from __future__ import annotations

import json
import re
import time
from pathlib import Path
from typing import Optional, Dict, Any, List

import requests
import pandas as pd

# ← 여기 새로 추가
from ..config import LOCAL_LLM_BASE_URL, LOCAL_LLM_MODEL


from .category_loader import load_category_master, MASTER_CACHE_FILE

# ===== Ollama 설정 =====
OLLAMA_BASE_URL = LOCAL_LLM_BASE_URL
OLLAMA_MODEL = LOCAL_LLM_MODEL

# ===== 1) 카테고리 마스터 로드 헬퍼 =====
_category_df: Optional[pd.DataFrame] = None


def get_category_master() -> pd.DataFrame:
    """pkl 또는 엑셀에서 카테고리 마스터를 로드하고, 메모리에 캐싱."""
    global _category_df
    if _category_df is not None:
        return _category_df

    def _log_cb(p: int, m: str):
        print(f"[cat_master] {p}% | {m}")

    df = load_category_master(progress_cb=_log_cb)
    print(f"카테고리 마스터 로드 완료: {len(df)}개")
    _category_df = df
    return df


# ===== 2) Ollama LLM 호출 래퍼 =====
class LLMError(RuntimeError):
    pass


def call_ollama_chat(system_prompt: str, user_prompt: str,
                     timeout: float = 300.0) -> str:
    """
    Ollama /api/chat 호출 래퍼.
    - system_prompt, user_prompt 를 넣고,
    - 최종 assistant 텍스트(content)만 반환.
    """
    url = f"{OLLAMA_BASE_URL}/api/chat"
    payload = {
        "model": OLLAMA_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "stream": False,
    }
    # --- 추가: llm 호출 전 진단 로그 ---
    try:
        sys_len = len(system_prompt or "")
        user_len = len(user_prompt or "")
        print(f"[LLM][req] url={url} model={OLLAMA_MODEL} timeout={timeout}")
        print(f"[LLM][req] prompt_len system={sys_len} user={user_len}")
    except Exception:
        pass
    
    # -------------------------
    # requests timeout은 (connect_timeout, read_timeout) 튜플을 쓰는 게 더 안전합니다.
    # - connect_timeout: 서버 접속 자체가 안 될 때 오래 멈추는 것 방지
    # - read_timeout: 응답이 늦어지는 경우 무한 대기 방지
    connect_timeout = min(10.0, float(timeout))        # 접속은 빠르게 실패시키기
    read_timeout = max(10.0, float(timeout))           # 전체 체감 타임아웃(기본 60s)
    req_timeout = (connect_timeout, read_timeout)

    last_err: Optional[Exception] = None
    max_attempts = 2  # 1회 재시도 (최소 침습)
    for attempt in range(1, max_attempts + 1):
        try:
            resp = requests.post(url, json=payload, timeout=req_timeout)
            resp.raise_for_status()
            break
        except requests.exceptions.Timeout as e:
            last_err = e
            if attempt >= max_attempts:
                raise LLMError(
                    f"LLM 호출 타임아웃 (connect={connect_timeout}s, read={read_timeout}s)"
                ) from e
            # 짧은 백오프 후 재시도
            time.sleep(0.6)
        except requests.RequestException as e:
            last_err = e
           # 네트워크/HTTP 계열은 1회 재시도 후 실패 처리
            if attempt >= max_attempts:
               raise LLMError(f"LLM HTTP 호출 실패: {e}") from e
            time.sleep(0.6)

    try:
        data = resp.json()
    except Exception as e:
        raise LLMError(f"LLM 응답 JSON 파싱 실패: {e} | raw={resp.text[:200]}") from e

    # Ollama /api/chat 응답 형식: { message: { role, content, ... }, ... }
    msg = data.get("message") or {}
    content = (msg.get("content") or "").strip()
    if not content:
        raise LLMError(f"LLM 응답이 비어 있습니다: {data}")

    return content


# ===== 3) 프롬프트 템플릿 =====

SYSTEM_PROMPT = """
너는 한국 온라인 쇼핑몰의 상품 카테고리를 추천하는 전문가다.

입력으로 상품명, 브랜드, 추가 설명이 주어지고,
카테고리 후보 목록(category_id, category_path 리스트)이 함께 주어진다.

너의 역할은:
1) 후보 목록 중에서 가장 적절한 카테고리 1개를 고르고
2) 그 이유를 짧게 설명하는 것이다.

반드시 JSON 형식으로만 출력해야 한다.
설명 문장은 JSON 바깥에 쓰지 말고, JSON 안의 "reason" 필드에 넣어라.

JSON 형식 (예시):

{
  "category_id": "12345",
  "category_path": "가전>주방가전>에어프라이어",
  "reason": "에어프라이어이며, 주방에서 사용하는 소형 가전이라 이 카테고리를 선택했습니다."
}

중요:
- category_id 는 반드시 **후보 목록에 존재하는 값**이어야 한다.
- category_path 또한 후보 목록에 있는 경로 중 하나를 그대로 사용해야 한다.
- 후보에 없는 id 나 path 를 새로 만들어내면 안 된다.

만약 적절한 카테고리를 고르기 어렵다면 다음처럼 응답하라:

{
  "category_id": null,
  "category_path": null,
  "reason": "카테고리를 판단하기 어려운 상품입니다. 추가 정보가 필요합니다."
}
""".strip()


USER_PROMPT_WITH_CANDIDATES_TEMPLATE = """
상품 정보는 다음과 같다.

- 상품명: {name}
- 브랜드: {brand}
- 추가 설명: {extra}

아래는 이 상품이 들어갈 수 있는 카테고리 후보 목록이다.
각 항목의 [] 안 값이 category_id 이고, 그 뒤가 category_path 이다.

{candidates_text}

위 후보 목록 중에서 딱 하나를 선택해서,
그 항목의 category_id, category_path 를 JSON 으로 반환해라.

반드시 위 후보 중 하나만 골라야 하며,
JSON의 category_id / category_path 값은 후보 목록의 값과 일치해야 한다.
추가 텍스트는 쓰지 말고, JSON 객체 한 개만 반환해라.
""".strip()


def _normalize_text(s: str) -> str:
    """소문자 + 공백 정리만 간단히 수행."""
    s = (s or "").strip()
    s = s.replace(">", " ")
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def _extract_keywords(name: str, brand: Optional[str], extra: Optional[str]) -> List[str]:
    """
    상품명/브랜드/추가 설명에서 2글자 이상 단어를 키워드로 추출.
    (아주 단순한 방식: 공백/특수문자 기준 split)
    """
    base = " ".join([x for x in [name, brand, extra] if x])
    base = base.replace(">", " ")
    # 한글/영문/숫자/공백만 남기고 나머지는 공백으로 치환
    base = re.sub(r"[^0-9a-zA-Z가-힣 ]+", " ", base)
    tokens = [t.strip() for t in base.split() if len(t.strip()) >= 2]

    # 중복 제거(순서 유지)
    seen = set()
    uniq = []
    for t in tokens:
        if t not in seen:
            seen.add(t)
            uniq.append(t)
    return uniq


def pick_candidate_categories(
    product_name: str,
    brand: Optional[str] = None,
    extra_text: Optional[str] = None,
    top_k: int = 30,
) -> List[Dict[str, str]]:
    """
    Step2: 전체 카테고리 마스터(1.6만개)에서
    상품명/브랜드/설명과 잘 겹치는 카테고리 TOP_k 를 뽑는다.

    간단 스코어링:
      - category_path 문자열에 키워드가 포함될수록 점수 +1
      - 스코어>0 인 것들 중 상위 top_k
      - 아무것도 없으면 앞에서부터 top_k 그냥 사용 (fallback)
    """
    df = get_category_master()
    if df.empty:
        return []

    keywords = _extract_keywords(product_name, brand, extra_text)
    if not keywords:
        # 키워드가 없으면 그냥 상위 몇 개만
        head = df.head(top_k)
        return [
            {
                "category_id": str(row["category_id"]),
                "category_path": str(row["category_path"]),
            }
            for _, row in head.iterrows()
        ]

    # 복사본 사용 (원본 df에 __score__ 컬럼 안 남기려고)
    df_tmp = df.copy()

    def score_row(row) -> int:
        path_text = _normalize_text(str(row["category_path"]))
        score = 0
        for kw in keywords:
            kw_norm = _normalize_text(kw)
            if kw_norm and kw_norm in path_text:
                score += 1
        return score

    df_tmp["__score__"] = df_tmp.apply(score_row, axis=1)
    scored = df_tmp[df_tmp["__score__"] > 0]

    if scored.empty:
        # 매칭되는 게 하나도 없으면, 그냥 앞에서부터 top_k
        head = df_tmp.head(top_k)
    else:
        head = scored.sort_values("__score__", ascending=False).head(top_k)

    candidates = [
        {
            "category_id": str(row["category_id"]),
            "category_path": str(row["category_path"]),
        }
        for _, row in head.iterrows()
    ]
    return candidates


def build_user_prompt_with_candidates(
    name: str,
    brand: Optional[str],
    extra: Optional[str],
    candidates: List[Dict[str, str]],
) -> str:
    """Step2용: 후보 카테고리 목록을 포함한 user 프롬프트 조립."""
    if not candidates:
        candidates_text = "(후보 카테고리가 없습니다. 이 경우, 적절한 카테고리를 고르기 어려울 수 있습니다.)"
    else:
        lines = []
        for i, c in enumerate(candidates, start=1):
            cid = c.get("category_id", "")
            cpath = c.get("category_path", "")
            lines.append(f"{i}. [{cid}] {cpath}")
        candidates_text = "\n".join(lines)

    return USER_PROMPT_WITH_CANDIDATES_TEMPLATE.format(
        name=name or "",
        brand=brand or "",
        extra=extra or "",
        candidates_text=candidates_text,
    )


# ===== 4) LLM 기반 카테고리 추천 엔트리 포인트 =====

def suggest_category_with_llm(
    product_name: str,
    brand: Optional[str] = None,
    extra_text: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Step2: 카테고리 마스터(pkl)에서 후보 TOP_k 뽑고,
    그 후보들 중에서만 LLM이 고르게 하는 구조.

    반환값 예:
    {
      "category_id": "12345" 또는 None,
      "category_path": "가전>주방가전>에어프라이어" 또는 None,
      "reason": "..."
    }
    """
    # 1) 카테고리 마스터 로드 (이미 캐싱됨)
    _ = get_category_master()

    # 2) 후보 카테고리 뽑기
    candidates = pick_candidate_categories(
        product_name=product_name,
        brand=brand,
        extra_text=extra_text,
        top_k=30,
    )

    # 후보가 전혀 없으면, 바로 실패 처리
    if not candidates:
        return {
            "category_id": None,
            "category_path": None,
            "reason": "카테고리 마스터에서 후보를 찾지 못했습니다.",
        }

    # LLM이 잘못된 path를 반환해도, 우리는 id 기준으로만 신뢰한다.
    cand_by_id = {c["category_id"]: c["category_path"] for c in candidates}

    user_prompt = build_user_prompt_with_candidates(
        product_name,
        brand,
        extra_text,
        candidates,
    )

    try:
        raw = call_ollama_chat(SYSTEM_PROMPT, user_prompt)
    except LLMError as e:
        # LLM 호출 자체가 실패한 경우
        return {
            "category_id": None,
            "category_path": None,
            "reason": f"LLM 호출 실패: {e}",
        }

    # LLM이 JSON 형식으로 잘 응답했다고 가정하고 파싱 시도
    # (혹시 앞뒤에 잡담이 붙어 있어도 첫 '{' ~ 마지막 '}' 구간만 파싱)
    try:
        start = raw.find("{")
        end = raw.rfind("}")
        if start == -1 or end == -1 or end <= start:
            raise ValueError(f"JSON 블록을 찾지 못했습니다: {raw[:200]}")

        json_str = raw[start : end + 1]
        obj = json.loads(json_str)

        raw_cat_id = obj.get("category_id")
        reason = obj.get("reason") or ""

        # null 이라면 판단 불가 처리
        if raw_cat_id is None:
            return {
                "category_id": None,
                "category_path": None,
                "reason": reason or "LLM이 카테고리를 판단할 수 없다고 응답했습니다.",
            }

        cat_id = str(raw_cat_id).strip()

        # LLM이 반환한 id가 후보 목록에 없으면, 실패 처리
        if cat_id not in cand_by_id:
            return {
                "category_id": None,
                "category_path": None,
                "reason": (
                    reason
                    + " (LLM이 후보 목록에 없는 category_id를 반환했습니다: "
                    + cat_id
                    + ")"
                ).strip(),
            }

        # path는 우리가 후보 목록에서 가져온 값으로 신뢰
        cat_path = cand_by_id[cat_id]

        return {
            "category_id": cat_id,
            "category_path": cat_path,
            "reason": reason,
        }

    except Exception as e:
        # JSON 형식이 어긋난 경우 raw 를 그대로 reason 에 남김
        return {
            "category_id": None,
            "category_path": None,
            "reason": f"LLM 응답 파싱 실패: {e} | raw={raw}",
        }
        
def suggest_category_with_candidates(
    product_name: str,
    brand: Optional[str],
    extra_text: Optional[str],
    candidates_df: Optional[pd.DataFrame],
) -> Dict[str, Any]:
    """
    candidates_df: category_id, category_path 컬럼을 가진 DataFrame
    - None 또는 empty면 기존 suggest_category_with_llm로 fallback
    - 아니면 후보 안에서만 고르게 LLM 프롬프트를 구성
    """

    # 0) 후보가 없으면 → 기존 전체 검색 함수로 위임 + 시간 측정
    if candidates_df is None or candidates_df.empty:
        start_ts = time.monotonic()
        result = suggest_category_with_llm(
            product_name=product_name,
            brand=brand,
            extra_text=extra_text,
        )
        elapsed = time.monotonic() - start_ts
        print(f"[LLM] 전체 검색 모드 호출 소요 시간: {elapsed:.2f}초")
        return result

    # 1) 후보 텍스트 구성
    candidates_text = "\n".join(
        f"- {row['category_id']}: {row['category_path']}"
        for _, row in candidates_df.iterrows()
    )

    # 2) system_prompt / user_prompt 구성 (여기서 항상 정의!)
    system_prompt = SYSTEM_PROMPT + """

너는 반드시 아래 '후보 카테고리 목록' 중에서만 하나를 골라야 한다.
후보 목록에 없는 카테고리는 새로 만들지 마라.
"""

    user_prompt = f"""
상품 정보:
- 상품명: {product_name}
- 브랜드: {brand or ""}
- 추가 설명: {extra_text or ""}

후보 카테고리 목록:
{candidates_text}

위 후보 중에서 가장 적절한 category_id 하나를 고르고,
해당 category_path와 reason을 JSON 한 개로만 출력해라.
"""

    # 3) LLM 호출 + 시간 측정
    try:
        start_ts = time.monotonic()
        raw = call_ollama_chat(system_prompt, user_prompt)
        elapsed = time.monotonic() - start_ts
        print(
            f"[LLM] 후보 제한 모드 호출 소요 시간: {elapsed:.2f}초 "
            f"(후보 수={len(candidates_df)})"
        )
    except LLMError as e:
        return {
            "category_id": None,
            "category_path": None,
            "reason": f"LLM 호출 실패(후보 제한 모드): {e}",
        }

    # 4) JSON 파싱 (기존 로직 그대로)
    try:
        start = raw.find("{")
        end = raw.rfind("}")
        if start == -1 or end == -1 or end <= start:
            raise ValueError(f"JSON 블록을 찾지 못했습니다: {raw[:200]}")

        json_str = raw[start : end + 1]
        obj = json.loads(json_str)

        cat_id = obj.get("category_id")
        cat_path = obj.get("category_path")
        reason = obj.get("reason") or ""

        return {
            "category_id": cat_id,
            "category_path": cat_path,
            "reason": reason,
        }
    except Exception as e:
        return {
            "category_id": None,
            "category_path": None,
            "reason": f"LLM 응답 파싱 실패(후보 제한 모드): {e} | raw={raw}",
        }




# ===== 5) 단독 실행용 테스트 =====

if __name__ == "__main__":
    # 1) 카테고리 마스터 로드
    get_category_master()

    # 2) 샘플 상품으로 테스트
    sample_name = "코스트코 스테인리스 양수냄비 24cm"
    sample_brand = "코스트코"
    sample_extra = "스테인리스 양수냄비, 24cm, 인덕션 가능, 주방용품"

    print("=== LLM 카테고리 추천 테스트 ===")
    result = suggest_category_with_llm(
        product_name=sample_name,
        brand=sample_brand,
        extra_text=sample_extra,
    )
    print(result)
