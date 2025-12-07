# cellon/category_ai/category_llm.py

from __future__ import annotations

import json
from pathlib import Path
from typing import Optional, Dict, Any

import requests

from .category_loader import load_category_master, MASTER_CACHE_FILE

# ===== Ollama 설정 =====
OLLAMA_BASE_URL = "http://localhost:11434"
# 설치된 모델 중 하나 선택: phi3:medium 또는 llama3:8b
OLLAMA_MODEL = "phi3:medium"


# ===== 1) 카테고리 마스터 로드 헬퍼 =====
_category_df = None


def get_category_master():
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
                     timeout: float = 180.0) -> str:
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

    try:
        resp = requests.post(url, json=payload, timeout=timeout)
        resp.raise_for_status()
    except requests.exceptions.Timeout as e:
        raise LLMError(f"LLM 호출 타임아웃 (timeout={timeout}s)") from e
    except requests.RequestException as e:
        raise LLMError(f"LLM HTTP 호출 실패: {e}") from e

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

입력으로 상품명, 브랜드, 추가 설명이 주어진다.
너의 역할은:
1) 상품이 들어갈 만한 대표 카테고리 id 와 전체 경로(category_path)를 하나 고르고
2) 그 이유를 짧게 설명하는 것이다.

반드시 JSON 형식으로만 출력해야 한다.
설명 문장을 JSON 바깥에 쓰지 말고, JSON 안의 "reason" 필드에 넣어라.

JSON 형식 (예시):

{
  "category_id": "12345",
  "category_path": "가전>주방가전>에어프라이어",
  "reason": "에어프라이어이며, 주방에서 사용하는 소형 가전이라 이 카테고리를 선택했습니다."
}

만약 적절한 카테고리를 고르기 어렵다면 다음처럼 응답하라:

{
  "category_id": null,
  "category_path": null,
  "reason": "카테고리를 판단하기 어려운 상품입니다. 추가 정보가 필요합니다."
}
""".strip()


USER_PROMPT_TEMPLATE = """
상품 정보는 다음과 같다.

- 상품명: {name}
- 브랜드: {brand}
- 추가 설명: {extra}

위 상품을 위한 카테고리 id, category_path, reason 을 JSON 한 개로만 출력해라.
추가 텍스트나 설명은 쓰지 말고, JSON 객체 딱 하나만 반환해라.
""".strip()


def build_user_prompt(
    name: str,
    brand: Optional[str] = None,
    extra: Optional[str] = None,
) -> str:
    """LLM에 넘길 user 프롬프트 조립."""
    brand = brand or ""
    extra = extra or ""
    return USER_PROMPT_TEMPLATE.format(name=name, brand=brand, extra=extra)


# ===== 4) LLM 기반 카테고리 추천 엔트리 포인트 =====

def suggest_category_with_llm(
    product_name: str,
    brand: Optional[str] = None,
    extra_text: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Step1: 순수 LLM 만으로 카테고리 추천 (카테고리 마스터는 단순 로드만).
    반환값 예:
    {
      "category_id": "12345" 또는 None,
      "category_path": "가전>주방가전>에어프라이어" 또는 None,
      "reason": "..."
    }
    """
    
    try:
        # ✅ LLM 호출 시간 측정 시작
        start_ts = time.monotonic()
        raw = call_ollama_chat(SYSTEM_PROMPT, user_prompt)
        elapsed = time.monotonic() - start_ts
        print(f"[LLM] suggest_category_with_llm 응답 시간: {elapsed:.2f}초 (후보 수={len(candidates)})")
        # ✅ 여기까지
    except LLMError as e:
        return {
            "category_id": None,
            "category_path": None,
            "reason": f"LLM 호출 실패: {e}",
        }
        
    # (필요하면 여기에서 get_category_master() 를 호출해서
    #  향후 Step2에서 유사도 검색 등에 활용할 수 있음)
    _ = get_category_master()

    user_prompt = build_user_prompt(product_name, brand, extra_text)

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

        cat_id = obj.get("category_id")
        cat_path = obj.get("category_path")
        reason = obj.get("reason") or ""

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
