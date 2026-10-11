import os
import json

BRAND_FILE = os.path.join(os.path.dirname(__file__), "brand_templates.json")
SELLER_MAP_FILE = os.path.join(os.path.dirname(__file__), "seller_mappings.json")

# 1. 브랜드사 템플릿 함수들
def load_brand_templates():
    if os.path.exists(BRAND_FILE):
        try:
            with open(BRAND_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "호무로": {
            "brand_name": "호무로",
            "columns": ["순번", "주문일시", "주문번호", "수령인", "연락처", "우편번호", "주소", "상세주소", "브랜드상품명", "옵션명", "수량", "배송메모", "제주도서여부", "이벤트색인"],
            "gift_rule": "add_row"
        },
        "기본_표준양식": {
            "brand_name": "기본_표준양식",
            "columns": ["순번", "셀러명", "주문일시", "주문번호", "수령인", "연락처", "우편번호", "주소", "브랜드상품명", "옵션명", "수량", "제주도서여부", "이벤트색인", "배송메모"],
            "gift_rule": "add_row"
        }
    }

def save_brand_templates(data):
    with open(BRAND_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def save_brand_template(brand_name, columns, gift_rule="add_row"):
    data = load_brand_templates()
    data[brand_name] = {
        "brand_name": brand_name,
        "columns": columns,
        "gift_rule": gift_rule
    }
    save_brand_templates(data)

# 2. 셀러 매핑 함수들
def load_seller_mappings():
    if os.path.exists(SELLER_MAP_FILE):
        try:
            with open(SELLER_MAP_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_seller_mappings(data):
    with open(SELLER_MAP_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def save_seller_mapping(seller_name, brand_name, column_mapping, product_mappings=None):
    data = load_seller_mappings()
    if product_mappings is None and seller_name in data:
        product_mappings = data[seller_name].get("product_mappings", {})
    elif product_mappings is None:
        product_mappings = {}

    data[seller_name] = {
        "seller_name": seller_name,
        "brand_name": brand_name,
        "column_mapping": column_mapping,
        "product_mappings": product_mappings
    }
    save_seller_mappings(data)

def get_seller_mapping(seller_name):
    data = load_seller_mappings()
    if seller_name in data:
        return data[seller_name]
    # 부분 일치 검색
    for key, val in data.items():
        if key in seller_name or seller_name in key:
            return val
    return None
