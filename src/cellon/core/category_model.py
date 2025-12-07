# cellon/core/category_model.py

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional

# ==============================
# 1. 마켓(목적지) Enum
# ==============================

class Marketplace(str, Enum):
    """
    최종 등록할 마켓 구분.
    - 쿠팡, 네이버 스토어, 기타 등
    """
    COUPANG = "coupang"         # 쿠팡 마켓
    NAVER_STORE = "naver_store" # 네이버 스토어
    ETC = "etc"                 # 기타 마켓

# ==============================
# 2. 카테고리 매칭 조건 모델
# ==============================

@dataclass
class CategoryCondition:
    """
    한 카테고리 규칙이 '언제' 적용될지를 정의하는 조건들.

    - required_keywords : 상품명/카테고리 힌트에 반드시 포함되어야 하는 키워드
    - optional_keywords : 포함되어 있으면 점수 가산
    - forbidden_keywords: 포함되어 있으면 탈락(또는 점수 크게 감점)

    - required_tags      : Product.tags 에 포함되어야 하는 태그
    - forbidden_tags     : Product.tags 에 있으면 안 되는 태그

    - attr_equals        : attributes[key] == value 여야 하는 조건
    - attr_contains      : attributes[key] 안에 특정 문자열/키워드가 포함되어야 하는 조건
                           (예: material 안에 "스테인리스" 포함)
    """
    required_keywords: List[str] = field(default_factory=list)      # 반드시 포함되어야 하는 키워드 리스트
    optional_keywords: List[str] = field(default_factory=list)      # 있으면 점수 가산되는 키워드 리스트
    forbidden_keywords: List[str] = field(default_factory=list)     # 있으면 탈락/감점되는 키워드 리스트

    required_tags: List[str] = field(default_factory=list)          # 반드시 포함되어야 하는 태그 리스트
    forbidden_tags: List[str] = field(default_factory=list)         # 있으면 안 되는 태그 리스트

    attr_equals: Dict[str, Any] = field(default_factory=dict)       # attributes[key] == value 조건
    attr_contains: Dict[str, List[str]] = field(default_factory=dict) # attributes[key]에 특정 값이 포함되어야 하는 조건

    def is_trivially_true(self) -> bool:
        """
        아무 조건도 없으면 '항상 참'인 규칙으로 볼 수 있음.
        (우선순위(priority)로만 제어할 때 사용 가능)
        """
        return (
            not self.required_keywords
            and not self.optional_keywords
            and not self.forbidden_keywords
            and not self.required_tags
            and not self.forbidden_tags
            and not self.attr_equals
            and not self.attr_contains
        )

# ==============================
# 3. 카테고리 규칙 모델
# ==============================

@dataclass
class CategoryRule:
    """
    '쿠팡 주방/냄비' 같은 한 카테고리를 매칭하기 위한 규칙 한 개.

    예:
    - rule_id      : "coupang_pot_24cm"
    - marketplace  : COUPANG
    - category_id  : "123456"
    - category_path: "주방용품 > 냄비/찜솥 > 양수냄비"
    - priority     : 10 (숫자가 작을수록 우선 적용)
    - conditions   : CategoryCondition(...)
    - default_fields: 쿠팡 엑셀/오픈API에 같이 들어가야 하는 기본값들
                      (예: delivery_type, tax_type 등)
    """
    rule_id: str                                         # 규칙 고유 ID
    marketplace: Marketplace                             # 적용 마켓
    category_id: str                                    # 마켓에서 사용하는 카테고리 ID
    category_path: str                                  # 사람이 보기 위한 한글 경로

    priority: int = 100                                 # 숫자 작을수록 우선순위 높음
    conditions: CategoryCondition = field(default_factory=CategoryCondition) # 적용 조건

    default_fields: Dict[str, Any] = field(default_factory=dict) # 기본 필드값들(엑셀/오픈API용)
    is_active: bool = True                              # 규칙 활성화 여부
    notes: Optional[str] = None                         # 메모/비고

# ==============================
# 4. 매칭 결과 모델
# ==============================

@dataclass
class MatchedCategory:
    """
    Product에 대해 CategoryRule을 적용한 '결과'.

    - rule_id      : 어떤 규칙과 매칭되었는지
    - marketplace  : 어떤 마켓용 카테고리인지
    - category_id  : 마켓에서 사용하는 카테고리 ID
    - category_path: 사람이 보기 위한 한글 경로
    - score        : 매칭 점수 (높을수록 더 잘 맞는다고 판단)
    - fields       : 실제 엑셀/쿠팡 API 등에 넘길 필드 세트
                     (CategoryRule.default_fields + Product 기반 가공 결과)
    """
    rule_id: str                                         # 적용된 규칙 ID
    marketplace: Marketplace                             # 적용된 마켓
    category_id: str                                    # 적용된 카테고리 ID
    category_path: str                                  # 적용된 카테고리 경로
    score: float                                        # 매칭 점수

    fields: Dict[str, Any] = field(default_factory=dict) # 실제 전달할 필드 세트

    def to_dict(self) -> Dict[str, Any]:
        """
        로그 저장이나 디버깅 용도로 dict로 변환.
        (엑셀/JSON 저장 등에 사용)
        """
        return {
            "rule_id": self.rule_id,                         # 규칙 ID
            "marketplace": self.marketplace.value,           # 마켓 이름(str)
            "category_id": self.category_id,                 # 카테고리 ID
            "category_path": self.category_path,             # 카테고리 경로
            "score": self.score,                             # 매칭 점수
            "fields": self.fields,                           # 필드 세트
        }
