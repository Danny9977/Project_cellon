# cellon/core/rules_loader.py

from __future__ import annotations  # 미래 버전 호환을 위한 import
import json
from pathlib import Path
from typing import List
from typing import List  # 타입 힌트용 import

from functools import lru_cache
from typing import Dict, Any

from cellon.config import BASE_DIR  # BASE_DIR가 config에 없다면 Path(__file__).resolve().parent.parent 정도로 조정

from .category_model import (
    CategoryRule,         # 카테고리 규칙 클래스
    CategoryCondition,    # 카테고리 조건 클래스
    Marketplace,          # 마켓 Enum
)

RULES_DIR = Path(__file__).resolve().parent.parent / "rules"
META_DIR = RULES_DIR / "meta"
COUPANG_DIR = RULES_DIR / "coupang"


# 프로젝트 루트 기준으로 rules 디렉토리 경로를 계산
BASE_DIR = Path(__file__).resolve().parents[2]  # .../Project_cellon
RULES_DIR = BASE_DIR / "rules"


def _condition_from_dict(data: dict) -> CategoryCondition:
    return CategoryCondition(
        required_keywords=data.get("required_keywords", []) or [],
        optional_keywords=data.get("optional_keywords", []) or [],
        forbidden_keywords=data.get("forbidden_keywords", []) or [],
        required_tags=data.get("required_tags", []) or [],
        forbidden_tags=data.get("forbidden_tags", []) or [],
        attr_equals=data.get("attr_equals", {}) or {},
        attr_contains=data.get("attr_contains", {}) or {},
    )


def _rule_from_dict(data: dict) -> CategoryRule:
    """
    JSON dict -> CategoryRule 객체 변환
    """
    conditions_data = data.get("conditions", {}) or {}
    conditions = _condition_from_dict(conditions_data)

    marketplace_str = data.get("marketplace", "etc")
    try:
        marketplace = Marketplace(marketplace_str)
    except ValueError:
        marketplace = Marketplace.ETC

    return CategoryRule(
        rule_id=data["rule_id"],
        marketplace=marketplace,
        category_id=data["category_id"],
        category_path=data.get("category_path", ""),
        priority=int(data.get("priority", 100)),
        conditions=conditions,
        default_fields=data.get("default_fields", {}) or {},
        is_active=bool(data.get("is_active", True)),
        notes=data.get("notes"),
    )


def load_rules_from_json(filename: str) -> List[CategoryRule]:
    """
    rules/ 디렉토리 아래의 JSON 파일에서 CategoryRule 리스트를 로딩.
    예:
        load_rules_from_json("coupang_demo_rules.json")
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

def _load_json(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


@lru_cache(maxsize=1)
def load_meta_kitchen_rules() -> Dict[str, Any]:
    """
    코스트코/도매매 → 메타 주방 카테고리 매핑 (coupang_kitchen.json)
    """
    path = META_DIR / "coupang_kitchen.json"
    return _load_json(path)


@lru_cache(maxsize=1)
def load_coupang_kitchen_rules() -> Dict[str, Any]:
    """
    메타 주방 카테고리 → 쿠팡 카테고리 ID 매핑 (kitchen_rules.json)
    """
    path = COUPANG_DIR / "kitchen_rules.json"
    return _load_json(path)

