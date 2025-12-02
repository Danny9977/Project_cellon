# cellon/core/manual_match_test.py

from __future__ import annotations

from ..product import Product, SourceDomain
from ..category_matcher import CategoryMatcher
from ..rules_loader import load_demo_rules


def build_demo_product() -> Product:
    """
    코스트코 엑셀/크롤링에서 가져왔다고 가정하는 '냄비' 상품 예시.
    실제로는 Product.from_costco_row(row)로 바꾸면 됩니다.
    """
    row = {
        "상품명": "코스트코 스테인리스 양수냄비 24cm",
        "브랜드": "코스트코",
        "카테고리": "주방용품>냄비",
        "용량": "24cm",
        "원본URL": "https://www.costco.co.kr/...",
        "원본ID": "123456",
        "_row_index": 10,
        "_sheet_name": "코스트코_냄비",
    }
    product = Product.from_costco_row(row)
    return product


def main() -> None:
    # 1) 데모 Product 생성
    product = build_demo_product()
    print("=== Product ===")
    print("display_name:", product.display_name)
    print("category_hint:", product.category_hint)
    print("attributes:", product.attributes)
    print()

    # 2) 데모 규칙 로딩
    rules = load_demo_rules()  # 내부적으로 coupang_demo_rules.json 로딩
    matcher = CategoryMatcher(rules)

    # 3) 매칭 실행
    results = matcher.match(product, top_k=3)

    print("=== Matched Categories ===")
    if not results:
        print("매칭된 카테고리가 없습니다.")
        return

    for idx, mc in enumerate(results, start=1):
        print(f"[{idx}] rule_id={mc.rule_id}")
        print(f"    marketplace  = {mc.marketplace.value}")
        print(f"    category_id  = {mc.category_id}")
        print(f"    category_path= {mc.category_path}")
        print(f"    score        = {mc.score}")
        print(f"    fields       = {mc.fields}")
        print()


if __name__ == "__main__":
    main()
