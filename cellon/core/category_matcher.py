# cellon/core/category_matcher.py
# LLM 한테 카테고리 매칭 코드를 만들고 있어서 당분간 이 코드 사용은 보류

from __future__ import annotations  # 미래 버전 호환을 위한 import

from typing import Iterable, List  # 타입 힌트용 import

from .product import Product  # Product 클래스 import
from .category_model import CategoryRule, CategoryCondition, MatchedCategory  # 카테고리 관련 모델 import

# 카테고리 매칭 엔진 클래스 정의
class CategoryMatcher:
    """
    Product + CategoryRule 리스트를 받아서
    최적의 카테고리를 찾아주는 매칭 엔진.

    사용 예:
        matcher = CategoryMatcher(rules)
        results = matcher.match(product, top_k=3)
    """

    def __init__(self, rules: Iterable[CategoryRule]):
        # 활성화된 규칙만 리스트로 저장, priority와 rule_id 기준으로 정렬
        self.rules: List[CategoryRule] = sorted(
            [r for r in rules if r.is_active],  # 활성화된 규칙만 필터링
            key=lambda r: (r.priority, r.rule_id),  # 우선순위, rule_id 기준 정렬
        )

    # ==========================
    # Public API
    # ==========================

    def match(self, product: Product, top_k: int = 3) -> List[MatchedCategory]:
        """
        Product에 대해 모든 규칙을 평가해서,
        점수가 높은 상위 top_k 개 카테고리를 반환.

        - score <= 0 인 규칙은 매칭 실패로 간주하고 버린다.
        """
        candidates: List[MatchedCategory] = []  # 후보 결과 리스트

        for rule in self.rules:  # 모든 규칙에 대해 반복
            score = self._score_rule(product, rule)  # 점수 계산
            if score <= 0:  # 점수가 0 이하이면 매칭 실패
                continue

            fields = self._build_fields(product, rule)  # 결과 필드 생성
            candidates.append(
                MatchedCategory(
                    rule_id=rule.rule_id,  # 규칙 ID
                    marketplace=rule.marketplace,  # 마켓 종류
                    category_id=rule.category_id,  # 카테고리 ID
                    category_path=rule.category_path,  # 카테고리 경로
                    score=score,  # 매칭 점수
                    fields=fields,  # 결과 필드
                )
            )

        # 점수 내림차순으로 정렬 후 상위 top_k 반환
        candidates.sort(key=lambda c: c.score, reverse=True)
        if top_k is not None and top_k > 0:
            return candidates[:top_k]  # top_k개만 반환
        return candidates  # top_k가 None이면 전체 반환

    # ==========================
    # 내부 헬퍼 메서드들
    # ==========================

    def _score_rule(self, product: Product, rule: CategoryRule) -> float:
        """
        하나의 rule에 대해 Product가 얼마나 잘 맞는지 점수를 계산.
        - 0 이하면 '매칭 실패'로 간주.
        """
        cond: CategoryCondition = rule.conditions  # 규칙의 조건 가져오기
        text = self._build_text_for_matching(product)  # 매칭용 텍스트 생성

        score = 0.0  # 초기 점수

        # 1) required_keywords: 하나라도 없으면 실패
        for kw in cond.required_keywords:
            if not self._contains(text, kw):  # 필수 키워드가 없으면
                return 0.0
            score += 10.0  # 필수 키워드는 가산

        # 2) optional_keywords: 있으면 가산
        for kw in cond.optional_keywords:
            if self._contains(text, kw):  # 선택 키워드가 있으면
                score += 3.0

        # 3) forbidden_keywords: 있으면 강한 패널티
        for kw in cond.forbidden_keywords:
            if self._contains(text, kw):  # 금지 키워드가 있으면
                return 0.0

        # 4) required_tags: Product.tags에 있어야 함
        for tag in cond.required_tags:
            if not self._tag_contains(product, tag):  # 필수 태그가 없으면
                return 0.0
            score += 5.0

        # 5) forbidden_tags: 있으면 실패
        for tag in cond.forbidden_tags:
            if self._tag_contains(product, tag):  # 금지 태그가 있으면
                return 0.0

        # 6) attr_equals: attributes[key] == value
        for key, value in cond.attr_equals.items():
            if product.get_attr(key) != value:  # 속성 값이 다르면
                return 0.0
            score += 4.0

        # 7) attr_contains: attributes[key]에 특정 키워드 포함
        for key, keywords in cond.attr_contains.items():
            attr_val = str(product.get_attr(key, "") or "").lower()  # 속성 값 가져오기
            for kw in keywords:
                if kw.lower() in attr_val:  # 키워드가 포함되어 있으면
                    score += 2.0
                else:
                    # 필수까지는 아니고, 없으면 점수만 안 올라감
                    pass

        # 우선순위(priority)가 낮을수록(숫자가 작을수록) 기본 보너스
        # (예: priority=10 → +5점, priority=100 → 거의 +0점)
        score += max(0.0, 20.0 - rule.priority * 0.2)

        return score  # 최종 점수 반환

    def _build_text_for_matching(self, product: Product) -> str:
        """
        키워드 매칭용으로 사용할 문자열을 구성.
        - raw_name, name, category_hint, tags 등을 합쳐서 한 문자열로 만든다.
        """
        parts = [
            product.raw_name or "",  # 원본 상품명
            product.name or "",      # 상품명
            product.category_hint or "",  # 카테고리 힌트
            " ".join(product.tags),  # 태그들
        ]
        return " ".join(p for p in parts if p).lower()  # 모두 합쳐서 소문자로 반환

    @staticmethod
    def _contains(text: str, keyword: str) -> bool:
        """
        keyword가 text 안에 포함되어 있는지(대소문자 무시).
        """
        if not keyword:
            return False
        return keyword.lower() in text  # 소문자로 비교

    @staticmethod
    def _tag_contains(product: Product, tag: str) -> bool:
        """
        Product.tags에 특정 태그가 포함되어 있는지.
        """
        tag = tag.strip().lower()  # 태그 소문자 변환
        return any(tag == t.strip().lower() for t in product.tags)  # 태그 리스트에 포함 여부

    def _build_fields(self, product: Product, rule: CategoryRule) -> dict:
        """
        최종 결과에 담길 fields dict 생성.
        - 기본은 CategoryRule.default_fields 복사
        - 필요하면 Product 기반으로 필드를 추가/가공
        (지금은 최소한으로만 구현, 나중에 확장)
        """
        fields = dict(rule.default_fields)  # 기본 필드 복사

        # 예: 상품명, 브랜드는 기본적으로 세팅 (마켓별로 나중에 조정 가능)
        fields.setdefault("product_name", product.display_name)
        if product.brand:
            fields.setdefault("brand", product.brand)

        # attributes 일부를 그대로 가져오는 것도 가능 (옵션)
        # 예: 용량(capacity), 재질(material) 등을 옮기는 등
        # 나중에 실제 엑셀/쿠팡 API 포맷에 맞춰 확장
        if product.get_attr("capacity"):
            fields.setdefault("capacity", product.get_attr("capacity"))
        if product.get_attr("material"):
            fields.setdefault("material", product.get_attr("material"))

        return fields  # 결과 필드
