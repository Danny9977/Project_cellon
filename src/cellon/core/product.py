# cellon/core/product.py

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional


# ==============================
# 1. 기본 Enum 정의
# ==============================

class SourceDomain(str, Enum):
    """
    원본 공급처 / 도메인 구분
    - costco : 코스트코
    - domebae : 도매매
    - ownerclan : 오너클랜
    - naver : 네이버 스마트스토어(또는 쇼핑)
    - etc : 기타 (임시)
    """
    COSTCO = "costco"
    DOMEBAE = "domebae"
    OWNERCLAN = "ownerclan"
    NAVER = "naver"
    ETC = "etc"


# ==============================
# 2. 옵션(색상/사이즈 등) 정의
# ==============================

@dataclass
class ProductOption:
    """
    상품 옵션(색상/사이즈/패키지 등)을 표현하는 모델.
    - 예: '블랙 / M', '대용량 / 2개세트' 등
    """
    name: str                            # 옵션 표시 이름 (예: "블랙 / M")
    attributes: Dict[str, Any] = field(default_factory=dict)
    """
    예:
    {
        "color": "블랙",
        "size": "M",
        "package": "2개입",
        "extra_price": 2000,   # 기본가 대비 추가 금액
    }
    """

    def get(self, key: str, default: Any = None) -> Any:
        return self.attributes.get(key, default)


# ==============================
# 3. Product(핵심 도메인 모델)
# ==============================

@dataclass
class Product:
    """
    '공급처에서 가져온 상품 1개'에 대한 표준 모델.

    - source_domain : 어떤 공급처에서 가져온 상품인지 (costco, domebae, ownerclan 등)
    - source_id     : 공급처에서 쓰는 상품 ID (없으면 None)
    - raw_name      : 원본 텍스트(공급처에 있는 그대로의 상품명)
    - name          : 가공/정제한 이름(브랜드 제거, 불필요 단어 제거 등) - 없으면 None
    - brand         : 브랜드명(있으면)
    - model_name    : 모델명/제품 시리즈명(있으면)
    - category_hint : '냄비', '의자', '바지', '샴푸'처럼 카테고리 유추에 도움이 되는 힌트
    - tags          : 검색에 도움이 되는 태그(키워드 리스트)
    - attributes    : 용량, 재질, 사이즈, 사용처 등 각종 속성
    - options       : 색상/사이즈/패키지 옵션들
    - meta          : 디버깅/추적용 메타 정보(원본 URL, 엑셀 row index 등)
    """
    # ====== 필수/주요 필드 ======
    source_domain: SourceDomain
    raw_name: str

    # ====== 식별자/추적 정보 ======
    source_id: Optional[str] = None      # 공급처 상품ID(상품코드 등)
    source_url: Optional[str] = None     # 공급처 상품 상세 페이지 URL

    # ====== 이름/브랜드/모델 ======
    name: Optional[str] = None           # 정제된 이름 (없으면 raw_name 사용)
    brand: Optional[str] = None
    model_name: Optional[str] = None

    # ====== 카테고리/검색 관련 ======
    category_hint: Optional[str] = None  # '냄비', '의자', '패딩바지' 등
    tags: List[str] = field(default_factory=list)

    # ====== 속성/옵션 ======
    attributes: Dict[str, Any] = field(default_factory=dict)
    options: List[ProductOption] = field(default_factory=list)

    # ====== 기타 메타 ======
    meta: Dict[str, Any] = field(default_factory=dict)

    # ------------------------------
    # 편의 메서드
    # ------------------------------

    @property
    def display_name(self) -> str:
        """
        UI / 로그 출력 등에 사용할 대표 상품명.
        - name 이 있으면 name
        - 없으면 raw_name
        """
        return self.name or self.raw_name

    def add_tag(self, *tags: str) -> None:
        """태그를 여러 개 한 번에 추가."""
        for t in tags:
            t = t.strip()
            if t and t not in self.tags:
                self.tags.append(t)

    def set_attr(self, key: str, value: Any) -> None:
        """attributes에 하나의 속성을 세팅."""
        self.attributes[key] = value

    def get_attr(self, key: str, default: Any = None) -> Any:
        """attributes에서 값 읽기."""
        return self.attributes.get(key, default)

    def add_option(self, option: ProductOption) -> None:
        """옵션 추가."""
        self.options.append(option)

    # ==============================
    # 4. 각 공급처별 편의 생성자 (예시)
    # ==============================

    @classmethod
    def from_costco_row(cls, row: Dict[str, Any]) -> "Product":
        """
        코스트코 크롤링/엑셀 row → Product 변환 템플릿.

        row 예시(가정):
        {
            "상품명": "코스트코 ○○양수냄비 24cm",
            "브랜드": "코스트코",
            "카테고리": "주방용품>냄비",
            "용량": "24cm",
            "원본URL": "https://www.costco.co.kr/...",
            "원본ID": "123456",
            ...
        }
        """
        raw_name = str(row.get("상품명", "")).strip()

        product = cls(
            source_domain=SourceDomain.COSTCO,
            raw_name=raw_name,
            source_id=str(row.get("원본ID") or "") or None,
            source_url=row.get("원본URL"),
        )

        # 브랜드/카테고리/기타 속성 매핑
        brand = row.get("브랜드")
        if brand:
            product.brand = str(brand).strip()

        category_text = row.get("카테고리")
        if category_text:
            product.category_hint = str(category_text).strip()
            product.add_tag(product.category_hint)

        # 예시: 용량, 재질 등의 속성을 attributes에 넣기
        if row.get("용량"):
            product.set_attr("capacity", str(row["용량"]).strip())
        if row.get("재질"):
            product.set_attr("material", str(row["재질"]).strip())

        # 메타 정보 (엑셀 row index, 시트명 등)
        if row.get("_row_index") is not None:
            product.meta["row_index"] = row["_row_index"]
        if row.get("_sheet_name"):
            product.meta["sheet_name"] = row["_sheet_name"]

        return product
