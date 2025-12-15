#src/cellon/crawlers/costco_crawler.py

from cellon.sellertool_excel import prepare_and_fill_sellertool


# ==========================
# 1) Product 생성
# ==========================
product = Product(
    source="costco",
    title=title,
    price=price,
    brand=brand,
    model=model,
    description=description,
    images=image_list,
    # ... 기존 코드 그대로
)

# ==========================
# 2) Coupang 카테고리 매칭
# ==========================
match = self.category_matcher.match(product)   # 예: category_matcher 객체가 있다면
product.coupang_category_id = match.category_id
product.coupang_category_path = match.category_path

# ==========================
# 3) 여기서 바로 Sellertool XLSM 생성!
# ==========================
try:
    xlsm_path = prepare_and_fill_sellertool(
        product=product,
        coupang_category_id=product.coupang_category_id,
        coupang_category_path=product.coupang_category_path,
        price=product.price,
        search_keywords=product.tags,      # 또는 크롤링 키워드 리스트
    )
    print(f"[셀러툴] 업로드 파일 생성됨: {xlsm_path}")

except Exception as e:
    print(f"[셀러툴 생성 실패] {e}")


