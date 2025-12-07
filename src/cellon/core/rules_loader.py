from __future__ import annotations

"""
cellon/core/rules_loader.py

- 기존 `load_meta_kitchen_rules()`, `load_coupang_kitchen_rules()`는 그대로 유지
- 추가로 마켓별(market) / 카테고리군(group) JSON을 자동 인식하는 헬퍼 함수들을 제공한다.

Directory layout (예시):

cellon/
  rules/
    meta/
      coupang_kitchen.json
      coupang_food.json
      ...
    coupang/
      kitchen_rules.json
      food_rules.json
      ...
    costco/
      kitchen.json
      ...
    domemae/
      kitchen.json
      ...
    owner/
      kitchen.json
      ...

각 market JSON의 포맷은 Danny님이 정의한대로:

{
  "__comment": {...},
  "categories": [
    {"name": "...", "eng": "...", "children": ["...", "..."]},
    ...
  ]
}
"""

import json
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping

# ----- 기본 경로 설정 -----

# 이 파일 위치: cellon/core/rules_loader.py
CELLON_DIR = Path(__file__).resolve().parent.parent
RULES_DIR = CELLON_DIR / "rules"
META_DIR = RULES_DIR / "meta"
COUPANG_DIR = RULES_DIR / "coupang"


# ----- 공통 JSON 로더 -----

def _load_json(path: Path) -> Any:
    """경로에 JSON 파일이 있으면 로드, 없거나 깨져 있으면 빈 dict 반환."""
    if not path.exists():
        return {}

    try:
        with path.open("r", encoding="utf-8") as f:
            text = f.read().strip()
            if not text:
                # 완전히 비어 있는 파일이면 그냥 빈 dict로 처리
                print(f"[rules_loader] Warning: empty JSON file: {path}")
                return {}
            return json.loads(text)
    except json.JSONDecodeError as e:
        print(f"[rules_loader] Warning: failed to parse JSON: {path} ({e})")
        # 깨진 파일이어도 앱이 죽지 않도록 빈 dict로 처리
        return {}



# ----- 1. 기존 룰(JSON list → CategoryRule) 로더 (유지) -----

from .category_model import CategoryCondition, CategoryRule, Marketplace  # pylint: disable=wrong-import-position


def _condition_from_dict(data: Mapping[str, Any]) -> CategoryCondition:
    return CategoryCondition(
        required_keywords=list(data.get("required_keywords", []) or []),
        optional_keywords=list(data.get("optional_keywords", []) or []),
        forbidden_keywords=list(data.get("forbidden_keywords", []) or []),
        required_tags=list(data.get("required_tags", []) or []),
        forbidden_tags=list(data.get("forbidden_tags", []) or []),
        attr_equals=dict(data.get("attr_equals", {}) or {}),
        attr_contains=dict(data.get("attr_contains", {}) or {}),
    )


def _rule_from_dict(data: Mapping[str, Any]) -> CategoryRule:
    """JSON dict → CategoryRule 객체 변환 (기존 coupang_demo_rules.json 용)."""
    conditions_data = data.get("conditions", {}) or {}
    conditions = _condition_from_dict(conditions_data)

    marketplace_str = data.get("marketplace", "etc")
    try:
        marketplace = Marketplace(marketplace_str)
    except ValueError:
        marketplace = Marketplace.ETC

    return CategoryRule(
        rule_id=str(data["rule_id"]),
        marketplace=marketplace,
        category_id=str(data["category_id"]),
        category_path=str(data.get("category_path", "")),
        priority=int(data.get("priority", 100)),
        conditions=conditions,
        default_fields=dict(data.get("default_fields", {}) or {}),
        is_active=bool(data.get("is_active", True)),
        notes=data.get("notes"),
    )


def load_rules_from_json(filename: str) -> List[CategoryRule]:
    """
    rules/ 디렉토리 아래의 JSON 파일에서 CategoryRule 리스트를 로딩.

    예: load_rules_from_json("coupang_demo_rules.json")
    """
    path = RULES_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"Rule JSON file not found: {path}")

    with path.open("r", encoding="utf-8") as f:
        raw = json.load(f)

    if not isinstance(raw, list):
        raise ValueError("Rule JSON root must be a list")

    rules: List[CategoryRule] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        rule = _rule_from_dict(item)
        rules.append(rule)
    return rules


# ----- 2. meta_kitchen / coupang_kitchen (기존) -----

@lru_cache(maxsize=1)
def load_meta_kitchen_rules() -> Dict[str, Any]:
    """코스트코/도매매 → 메타 주방 카테고리 매핑 (coupang_kitchen.json)."""
    path = META_DIR / "coupang_kitchen.json"
    data = _load_json(path)
    if not isinstance(data, dict):
        return {}
    return data


@lru_cache(maxsize=1)
def load_coupang_kitchen_rules() -> Dict[str, Any]:
    """메타 주방 카테고리 → 쿠팡 카테고리 ID 매핑 (kitchen_rules.json)."""
    path = COUPANG_DIR / "kitchen_rules.json"
    data = _load_json(path)
    if not isinstance(data, dict):
        return {}
    return data


# ----- 3. 확장: group 단위 meta / coupang 룰 자동 인식 -----

@lru_cache(maxsize=None)
def list_available_groups() -> List[str]:
    """
    meta/ 아래 coupang_*.json, coupang/ 아래 *_rules.json 을 스캔해서
    사용 가능한 카테고리군(group) 이름 목록을 반환.

    예: ["kitchen", "food", "beauty", ...]
    """
    groups: set[str] = set()

    if META_DIR.exists():
        for p in META_DIR.glob("coupang_*.json"):
            groups.add(p.stem.replace("coupang_", ""))

    if COUPANG_DIR.exists():
        for p in COUPANG_DIR.glob("*_rules.json"):
            groups.add(p.stem.replace("_rules", ""))

    return sorted(groups)


@lru_cache(maxsize=None)
def load_meta_rules(group: str) -> Dict[str, Any]:
    """
    특정 group에 대한 meta 룰 로딩.

    - group="kitchen"  → meta/coupang_kitchen.json
    - group="food"     → meta/coupang_food.json
    - group="beauty"   → meta/coupang_beauty.json
    """
    path = META_DIR / f"coupang_{group}.json"
    data = _load_json(path)
    if not isinstance(data, dict):
        return {}
    return data

# ----- 3-1. coupang 룰 업데이트 헬퍼 -----
@lru_cache(maxsize=None)
def load_coupang_rules(group: str) -> Dict[str, Any]:
    """ 특정 group에 대한 coupang 룰 로딩.
    - group="kitchen" → coupang/kitchen_rules.json
    - group="food"    → coupang/food_rules.json
    """
    path = COUPANG_DIR / f"{group}_rules.json"
    data = _load_json(path)
    if not isinstance(data, dict):
        return {}
    return data


def upsert_strong_name_rule(
    group: str,
    target_category_id: str | int,
    keywords: list[str],
    reason: str | None = None,
) -> dict:
    """
    - rules/coupang/{group}_rules.json 의 __strong_name_rules__ 를 업데이트.
    - 같은 category_id 에 대해서는 기존 rule 의 keywords 에만 합쳐 넣고 중복 제거.
    - 파일이 없으면 새로 생성.
    - 업데이트 후 load_coupang_rules 캐시를 비운다.
    """
    target_id_str = str(target_category_id).strip()
    # 키워드 정리
    cleaned: list[str] = []
    for kw in (keywords or []):
        kw = (kw or "").strip()
        if not kw:
            continue
        if kw not in cleaned:
            cleaned.append(kw)

    if not cleaned:
        # 넣을 키워드가 없으면 아무 것도 하지 않음
        return {}

    path = COUPANG_DIR / f"{group}_rules.json"
    path.parent.mkdir(parents=True, exist_ok=True)

    data = _load_json(path)
    if not isinstance(data, dict):
        data = {}

    rules_list = data.get("__strong_name_rules__", [])
    if not isinstance(rules_list, list):
        rules_list = []

    # 동일 category_id rule 찾기
    existing = None
    for rule in rules_list:
        if str(rule.get("target_category_id")) == target_id_str:
            existing = rule
            break

    if existing is not None:
        # 기존 keywords 에 합쳐 넣고 중복 제거
        existing_kw = [
            (str(k) or "").strip()
            for k in (existing.get("keywords") or [])
            if (str(k) or "").strip()
        ]
        for kw in cleaned:
            if kw not in existing_kw:
                existing_kw.append(kw)
        existing["keywords"] = existing_kw
        if reason:
            existing["reason"] = reason
    else:
        new_rule = {
            "keywords": cleaned,
            "target_category_id": target_id_str,
        }
        if reason:
            new_rule["reason"] = reason
        rules_list.append(new_rule)

    data["__strong_name_rules__"] = rules_list

    # JSON 저장
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # 캐시 무효화 (다음부터는 새 JSON 사용)
    load_coupang_rules.cache_clear()

    return data


# ----- 4. 마켓별 카테고리 트리(JSON) 자동 인식 -----

def _iter_market_dirs() -> Iterable[Path]:
    """rules/ 아래에서 meta, coupang 을 제외한 market 디렉토리들을 순회."""
    if not RULES_DIR.exists():
        return []
    for d in RULES_DIR.iterdir():
        if not d.is_dir():
            continue
        if d.name in {"meta", "coupang"}:
            continue
        yield d


@lru_cache(maxsize=None)
def list_markets() -> List[str]:
    """
    rules/ 아래 market 디렉토리 이름 목록.

    예: ["costco", "domemae", "owner"]
    """
    return sorted(d.name for d in _iter_market_dirs())


@lru_cache(maxsize=None)
def list_market_groups(market: str) -> List[str]:
    """
    특정 market 아래에 존재하는 group(json 파일 이름) 목록.

    예: rules/costco/kitchen.json → group "kitchen"
    """
    mdir = RULES_DIR / market
    if not mdir.is_dir():
        return []
    return sorted(p.stem for p in mdir.glob("*.json"))


@lru_cache(maxsize=None)
def load_market_group_json(market: str, group: str) -> Dict[str, Any]:
    """
    rules/<market>/<group>.json 을 로드.

    예:
      load_market_group_json("costco", "kitchen")
      load_market_group_json("domemae", "kitchen")
    """
    path = RULES_DIR / market / f"{group}.json"
    data = _load_json(path)
    if not isinstance(data, dict):
        return {}
    return data


@lru_cache(maxsize=1)
def load_all_market_groups() -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    모든 market / group JSON 을 한 번에 로드.

    반환 구조 예:

    {
      "costco": {
        "kitchen": {...},
        "food": {...},
      },
      "domemae": {
        "kitchen": {...}
      }
    }
    """
    result: Dict[str, Dict[str, Dict[str, Any]]] = {}

    for market in list_markets():
        groups: Dict[str, Dict[str, Any]] = {}
        for group in list_market_groups(market):
            groups[group] = load_market_group_json(market, group)
        result[market] = groups

    return result
