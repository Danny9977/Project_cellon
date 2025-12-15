# manual_test_seller_excel.py

from cellon.core.product import Product
from cellon.sellertool_excel import prepare_and_fill_sellertool

# 1) Product 생성
product = Product(
    source_domain="costco",
    raw_name="테스트 상품명 raw",
    name="테스트 상품명",
    attributes={"price": 12900, "brand": "테스트브랜드"},
    tags=["코스트코", "테스트"]
)

# 2) 엑셀 템플릿 카테고리 정보
category_id = "80289"
category_path = "주방용품>취사도구>냄비>양수냄비"

# 3) Excel 생성 실행
path = prepare_and_fill_sellertool(
    product=product,
    coupang_category_id=category_id,
    coupang_category_path=category_path,
    price=product.attributes.get("price"),
    search_keywords=product.tags,
)

print("생성된 파일:", path)
