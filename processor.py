import re
import io
import json
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 표준 필드 정의
STANDARD_FIELDS = {
    "order_time": "주문일시/결제일시",
    "order_id": "주문번호/관리코드",
    "buyer_name": "주문자(구매자)명",
    "buyer_phone": "주문자 연락처",
    "receiver_name": "수령인(받는분)명",
    "receiver_phone": "수령인 연락처",
    "zipcode": "우편번호",
    "address": "기본 배송주소",
    "address_detail": "상세 배송주소",
    "product_name": "상품명/품목",
    "option_name": "옵션명/선택사항",
    "quantity": "수량",
    "price": "결제금액/단가",
    "shipping_memo": "배송메모/요청사항",
    "seller_channel": "판매처/쇼핑몰명"
}

STANDARD_COLUMN_ALIASES = {
    "order_time": ["주문일시", "결제일시", "주문일자", "결제일", "주문시간", "order_date", "date", "주문일"],
    "order_id": ["주문번호", "관리번호", "주문id", "order_id", "주문코드", "쇼핑몰 주문번호", "쇼핑몰주문번호"],
    "buyer_name": ["주문자명", "구매자", "주문자", "고객명", "구매자명"],
    "buyer_phone": ["주문자연락처", "구매자전화번호", "주문자핸드폰", "구매자연락처"],
    "receiver_name": ["수령인", "받는분", "받으시는분", "수령인명", "수취인명", "받는사람", "수령자"],
    "receiver_phone": ["수령인연락처", "받는분연락처", "받으시는분전화번호", "수취인연락처", "전화번호", "휴대폰", "수령자전화번호", "수령자핸드폰"],
    "zipcode": ["우편번호", "우편", "zipcode", "postcode"],
    "address": ["주소", "기본주소", "배송지", "배송주소", "배송지주소", "수령인주소", "수령자주소"],
    "address_detail": ["상세주소", "배송지상세", "상세배송주소", "배송지상세주소"],
    "product_name": ["상품명", "주문상품", "품목명", "품목", "상품", "product_name"],
    "option_name": ["옵션", "옵션정보", "선택옵션", "옵션명", "규격"],
    "quantity": ["수량", "주문수량", "수량(개)", "qty"],
    "price": ["결제금액", "판매가", "결제액", "단가", "주문금액", "금액"],
    "shipping_memo": ["배송메세지", "배송메모", "배송요청사항", "메모", "배송요청", "요청사항"],
    "seller_channel": ["판매처", "쇼핑몰", "쇼핑몰명", "판매처명", "셀러명", "채널", "채널명", "판매채널", "몰명", "store", "channel"]
}

def detect_column_mapping(df_columns):
    """셀러 발주서 컬럼 초기 추천 (단순 참고용, 사용자가 직접 확인/수정 가능)"""
    col_map = {}
    normalized_cols = {str(c).strip().replace(" ", "").lower(): str(c) for c in df_columns}
    
    for std_key, aliases in STANDARD_COLUMN_ALIASES.items():
        found = False
        for alias in aliases:
            clean_alias = alias.replace(" ", "").lower()
            if clean_alias in normalized_cols:
                col_map[std_key] = normalized_cols[clean_alias]
                found = True
                break
        if not found:
            for norm_c, orig_c in normalized_cols.items():
                for alias in aliases:
                    if alias.replace(" ", "").lower() in norm_c:
                        col_map[std_key] = orig_c
                        found = True
                        break
                if found:
                    break
        if not found:
            col_map[std_key] = "(선택안함)"
    return col_map

def is_jeju_island(address, zipcode=""):
    """제주 및 도서산간 지역 여부 판별"""
    if pd.isna(address):
        address = ""
    addr_str = str(address).strip()
    
    jeju_keywords = ["제주", "서귀포", "울릉", "신안군", "옹진군", "백령", "연평", "거문도", "완도군 노화", "청산면"]
    for kw in jeju_keywords:
        if kw in addr_str:
            return True
            
    if zipcode and not pd.isna(zipcode):
        zip_clean = re.sub(r"[^0-9]", "", str(zipcode))
        if len(zip_clean) >= 5:
            try:
                zip_num = int(zip_clean[:5])
                if 63000 <= zip_num <= 63644:
                    return True
            except ValueError:
                pass
    return False

def extract_unique_products(df, col_map):
    """셀러 데이터프레임에서 고유한 (상품명, 옵션명) 목록 추출"""
    prod_col = col_map.get("product_name")
    opt_col = col_map.get("option_name")
    
    if not prod_col or prod_col == "(선택안함)" or prod_col not in df.columns:
        return []
    
    items = []
    has_opt = opt_col and opt_col != "(선택안함)" and opt_col in df.columns
    
    grouped = df[[prod_col] + ([opt_col] if has_opt else [])].drop_duplicates()
    for _, row in grouped.iterrows():
        p_val = str(row[prod_col]).strip() if not pd.isna(row[prod_col]) else ""
        o_val = str(row[opt_col]).strip() if has_opt and not pd.isna(row[opt_col]) else ""
        if p_val and p_val != "nan":
            items.append((p_val, o_val))
    return items

def process_seller_dataframe(df, seller_name, col_map, product_mapping_dict=None, custom_order_date=None, force_custom_date=False, custom_seller_name=None):
    """
    사용자가 지정한 컬럼 매핑 및 상품 매핑 사전을 바탕으로 데이터 정규화
    product_mapping_dict: { "(seller_prod, seller_opt)": {"brand_product": "...", "brand_option": "..."} }
    """
    norm_data = []

    # 🏪 판매처(쇼핑몰명) 전처리
    custom_seller_clean = str(custom_seller_name).strip() if custom_seller_name else ""
    channel_col = col_map.get("seller_channel", "")
    has_channel_col = bool(channel_col and channel_col != "(선택안함)" and channel_col in df.columns)

    if has_channel_col:
        s_channels = df[channel_col].astype(str).replace(["", "nan", "None", "(없음)"], np.nan).ffill()
    else:
        s_channels = None
    
    def get_val(row, field_key, default=""):
        col_name = col_map.get(field_key, "")
        if col_name and col_name != "(선택안함)" and col_name in row:
            val = row[col_name]
            return "" if pd.isna(val) else str(val).strip()
        return default

    for idx, row in df.iterrows():
        order_time_raw = get_val(row, "order_time")
        order_id = get_val(row, "order_id", f"ORD_{idx+1}")
        buyer_name = get_val(row, "buyer_name")
        buyer_phone = get_val(row, "buyer_phone")
        receiver_name = get_val(row, "receiver_name", buyer_name)
        receiver_phone = get_val(row, "receiver_phone", buyer_phone)
        zipcode = get_val(row, "zipcode")
        address = get_val(row, "address")
        address_detail = get_val(row, "address_detail")
        raw_product = get_val(row, "product_name")
        raw_option = get_val(row, "option_name")
        shipping_memo = get_val(row, "shipping_memo")

        # 판매처 결정
        if custom_seller_clean:
            row_seller = custom_seller_clean
        elif s_channels is not None and pd.notna(s_channels.iloc[idx]):
            row_seller = str(s_channels.iloc[idx]).strip()
        else:
            row_seller = seller_name

        # 수량
        qty_raw = get_val(row, "quantity", "1")
        try:
            qty = int(float(qty_raw.replace(",", ""))) if qty_raw else 1
        except Exception:
            qty = 1

        # 금액
        price_raw = get_val(row, "price", "0")
        try:
            price = int(float(price_raw.replace(",", ""))) if price_raw else 0
        except Exception:
            price = 0

        # 날짜 파싱 & 커스텀 일자 적용
        parsed_time = None
        if force_custom_date and custom_order_date:
            if hasattr(custom_order_date, "strftime"):
                order_date = custom_order_date.strftime("%Y-%m-%d")
            else:
                order_date = str(custom_order_date).strip()
            order_time_str = f"{order_date} 00:00:00"
            try:
                parsed_time = pd.Timestamp(order_date)
            except Exception:
                parsed_time = pd.Timestamp.now()
        else:
            if order_time_raw:
                try:
                    parsed_time = pd.to_datetime(order_time_raw)
                except Exception:
                    pass

            if parsed_time is None or pd.isna(parsed_time):
                if custom_order_date:
                    if hasattr(custom_order_date, "strftime"):
                        order_date = custom_order_date.strftime("%Y-%m-%d")
                    else:
                        order_date = str(custom_order_date).strip()
                    try:
                        parsed_time = pd.Timestamp(order_date)
                    except Exception:
                        parsed_time = pd.Timestamp.now()
                else:
                    parsed_time = pd.Timestamp.now()
                    order_date = parsed_time.strftime("%Y-%m-%d")
            else:
                order_date = parsed_time.strftime("%Y-%m-%d")
            order_time_str = parsed_time.strftime("%Y-%m-%d %H:%M:%S") if order_time_raw else f"{order_date} 00:00:00"

        # 제주/도서산간
        full_addr = f"{address} {address_detail}".strip()
        is_jeju = is_jeju_island(full_addr, zipcode)

        # 상품명 매핑 적용
        std_product = raw_product
        std_option = raw_option
        if product_mapping_dict:
            key_exact = f"{raw_product}||{raw_option}"
            key_prod_only = f"{raw_product}||"
            if key_exact in product_mapping_dict:
                std_product = product_mapping_dict[key_exact].get("brand_product", raw_product)
                std_option = product_mapping_dict[key_exact].get("brand_option", raw_option)
            elif key_prod_only in product_mapping_dict:
                std_product = product_mapping_dict[key_prod_only].get("brand_product", raw_product)
                std_option = product_mapping_dict[key_prod_only].get("brand_option", raw_option)

        norm_data.append({
            "seller_name": row_seller,
            "order_time_obj": parsed_time,
            "order_time": order_time_str,
            "order_date": order_date,
            "order_id": order_id,
            "buyer_name": buyer_name,
            "buyer_phone": buyer_phone,
            "receiver_name": receiver_name,
            "receiver_phone": receiver_phone,
            "zipcode": zipcode,
            "address": address,
            "address_detail": address_detail,
            "full_address": full_addr,
            "raw_product": raw_product,
            "raw_option": raw_option,
            "product_name": std_product,
            "option_name": std_option,
            "quantity": qty,
            "price": price,
            "shipping_memo": shipping_memo,
            "is_jeju": is_jeju,
            "event_tags": [],
            "gifts_to_add": []
        })

    return pd.DataFrame(norm_data)

def apply_event_rules(df, event_config):
    """이벤트 규칙 적용 및 사은품 객체 생성"""
    if df.empty:
        return df

    # 1. 선착순 이벤트
    fcfs_cfg = event_config.get("first_come_first_served", {})
    if fcfs_cfg.get("enabled", False):
        limit = fcfs_cfg.get("limit", 3)
        gift_prod = fcfs_cfg.get("gift_product_name", "[증정] 선착순 사은품")
        gift_opt = fcfs_cfg.get("gift_option_name", "증정용")
        gift_qty = fcfs_cfg.get("gift_quantity", 1)
        action_type = fcfs_cfg.get("action_type", "add_row")

        sorted_df = df.sort_values(by="order_time_obj")
        unique_orders = sorted_df["order_id"].drop_duplicates().tolist()
        winner_order_ids = set(unique_orders[:limit])

        for idx, row in df.iterrows():
            if row["order_id"] in winner_order_ids:
                row["event_tags"].append(f"[선착순{limit}명]")
                if action_type == "add_row":
                    row["gifts_to_add"].append({
                        "product_name": gift_prod,
                        "option_name": gift_opt,
                        "quantity": gift_qty,
                        "reason": f"선착순 {limit}명 사은품"
                    })

    # 2. 금액 조건 이벤트
    amt_cfg = event_config.get("min_amount", {})
    if amt_cfg.get("enabled", False):
        min_amt = amt_cfg.get("min_amount", 50000)
        gift_prod = amt_cfg.get("gift_product_name", "[증정] 금액달성 사은품")
        gift_opt = amt_cfg.get("gift_option_name", "증정용")
        gift_qty = amt_cfg.get("gift_quantity", 1)
        
        order_totals = df.groupby("order_id")["price"].sum().to_dict()
        for idx, row in df.iterrows():
            if order_totals.get(row["order_id"], 0) >= min_amt:
                tag = f"[{min_amt:,}원이상]"
                if tag not in row["event_tags"]:
                    row["event_tags"].append(tag)
                    already_has = any(g["reason"] == "금액달성 사은품" for g in row["gifts_to_add"])
                    if not already_has:
                        row["gifts_to_add"].append({
                            "product_name": gift_prod,
                            "option_name": gift_opt,
                            "quantity": gift_qty,
                            "reason": "금액달성 사은품"
                        })

    # 3. 특정 품목 구매 사은품
    spec_cfg = event_config.get("special_product", {})
    if spec_cfg.get("enabled", False):
        target_kw = spec_cfg.get("keyword", "")
        gift_prod = spec_cfg.get("gift_product_name", "[사은품] 특정품목 증정")
        gift_opt = spec_cfg.get("gift_option_name", "단품")
        gift_qty = spec_cfg.get("gift_quantity", 1)

        if target_kw:
            for idx, row in df.iterrows():
                if target_kw in f"{row['product_name']} {row['option_name']} {row['raw_product']}":
                    row["event_tags"].append(f"[{target_kw}구매]")
                    row["gifts_to_add"].append({
                        "product_name": gift_prod,
                        "option_name": gift_opt,
                        "quantity": gift_qty,
                        "reason": f"{target_kw} 구매 사은품"
                    })

    return df

def build_brand_order_sheet(df, brand_template):
    """
    브랜드사가 지정한 컬럼 양식대로 최종 발주서 생성
    """
    brand_cols = brand_template.get("columns", [
        "순번", "주문일시", "주문번호", "수령인", "연락처", "우편번호", "주소", "상세주소",
        "브랜드상품명", "옵션명", "수량", "배송메모", "제주도서여부", "이벤트색인"
    ])
    
    rows = []
    seq = 1

    for idx, r in df.iterrows():
        event_str = " / ".join(r["event_tags"]) if r["event_tags"] else ""
        
        # 기본 매핑 필드 풀
        field_pool = {
            "순번": seq,
            "셀러명": r["seller_name"],
            "쇼핑몰명": r["seller_name"],
            "판매처": r["seller_name"],
            "판매처명": r["seller_name"],
            "날짜": r["order_date"],
            "주문일자": r["order_date"],
            "발주일자": r["order_date"],
            "주문일시": r["order_time"],
            "주문번호": r["order_id"],
            "구매자": r["buyer_name"],
            "구매자연락처": r["buyer_phone"],
            "수령인": r["receiver_name"],
            "받는분": r["receiver_name"],
            "수령자": r["receiver_name"],
            "연락처": r["receiver_phone"],
            "전화번호": r["receiver_phone"],
            "우편번호": r["zipcode"],
            "주소": r["address"],
            "상세주소": r["address_detail"],
            "전체주소": r["full_address"],
            "브랜드상품명": r["product_name"],
            "상품명": r["product_name"],
            "옵션명": r["option_name"],
            "옵션": r["option_name"],
            "수량": r["quantity"],
            "결제금액": r["price"],
            "배송메모": r["shipping_memo"],
            "제주도서여부": "제주/도서산간" if r["is_jeju"] else "일반",
            "이벤트색인": event_str,
            "비고": event_str
        }

        row_dict = {}
        for col in brand_cols:
            row_dict[col] = field_pool.get(col, "")
        rows.append(row_dict)
        seq += 1

        # 사은품 행 추가
        if r.get("gifts_to_add"):
            for gift in r["gifts_to_add"]:
                gift_pool = field_pool.copy()
                gift_pool["순번"] = seq
                gift_pool["브랜드상품명"] = gift["product_name"]
                gift_pool["상품명"] = gift["product_name"]
                gift_pool["옵션명"] = gift["option_name"]
                gift_pool["옵션"] = gift["option_name"]
                gift_pool["수량"] = gift["quantity"]
                gift_pool["이벤트색인"] = f"[사은품] {gift['reason']}"
                gift_pool["비고"] = f"[사은품] {gift['reason']}"
                gift_pool["배송메모"] = "본품합포장 출고"

                g_row = {}
                for col in brand_cols:
                    g_row[col] = gift_pool.get(col, "")
                rows.append(g_row)
                seq += 1

    return pd.DataFrame(rows)

def build_daily_summary_sheets(df, shipping_config):
    """일자별 품목 수량표 및 배송비 정산표"""
    free_threshold = shipping_config.get("free_shipping_threshold", 50000)
    base_shipping_fee = shipping_config.get("base_shipping_fee", 3000)
    jeju_extra_fee = shipping_config.get("jeju_extra_fee", 3000)

    item_df = df.copy()
    item_df["품목(옵션)"] = item_df["product_name"] + " (" + item_df["option_name"] + ")"
    daily_qty_pivot = pd.pivot_table(
        item_df,
        index="order_date",
        columns="품목(옵션)",
        values="quantity",
        aggfunc="sum",
        fill_value=0
    )
    daily_qty_pivot["일자별_합계수량"] = daily_qty_pivot.sum(axis=1)
    daily_qty_pivot.reset_index(inplace=True)
    daily_qty_pivot.rename(columns={"order_date": "주문일자"}, inplace=True)

    order_groups = df.groupby(["order_date", "order_id"]).agg({
        "price": "sum",
        "is_jeju": "first",
        "quantity": "sum"
    }).reset_index()

    order_groups["기본배송비"] = order_groups["price"].apply(
        lambda p: 0 if p >= free_threshold else base_shipping_fee
    )
    order_groups["제주추가비"] = order_groups["is_jeju"].apply(
        lambda j: jeju_extra_fee if j else 0
    )
    order_groups["총배송비"] = order_groups["기본배송비"] + order_groups["제주추가비"]

    daily_shipping_summary = order_groups.groupby("order_date").agg(
        총_주문건수=("order_id", "count"),
        총_구매수량=("quantity", "sum"),
        무료배송_건수=("기본배송비", lambda s: (s == 0).sum()),
        유료배송_건수=("기본배송비", lambda s: (s > 0).sum()),
        제주도서_건수=("is_jeju", lambda s: s.sum()),
        기본배송비_합계=("기본배송비", "sum"),
        제주배송비_합계=("제주추가비", "sum"),
        총_배송비_합계=("총배송비", "sum")
    ).reset_index()
    daily_shipping_summary.rename(columns={"order_date": "주문일자"}, inplace=True)

    return daily_qty_pivot, daily_shipping_summary

def export_styled_excel(brand_df, daily_qty_df, daily_shipping_df, master_raw_df):
    """스타일링된 통합 엑셀 파일 생성"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        brand_df.to_excel(writer, sheet_name="브랜드사_발주서", index=False)
        daily_qty_df.to_excel(writer, sheet_name="일자별_품목수량표", index=False)
        daily_shipping_df.to_excel(writer, sheet_name="일자별_배송정산표", index=False)
        master_raw_df.to_excel(writer, sheet_name="셀러통합_판매DB", index=False)

    output.seek(0)
    wb = openpyxl.load_workbook(output)

    header_fill = PatternFill(start_color="1E3D59", end_color="1E3D59", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="맑은 고딕", size=9)
    highlight_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    gift_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        ws.views.sheetView[0].showGridLines = True
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws.row_dimensions[1].height = 28

        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 20
            
            is_gift_row = False
            is_event_tagged = False
            if sheetname == "브랜드사_발주서":
                row_str = " ".join(str(ws.cell(row=row_idx, column=c).value or "") for c in range(1, ws.max_column + 1))
                if "[사은품]" in row_str:
                    is_gift_row = True
                elif any(kw in row_str for kw in ["[선착순", "원이상", "구매]"]):
                    is_event_tagged = True

            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

                if is_gift_row:
                    cell.fill = gift_fill
                elif is_event_tagged:
                    cell.fill = highlight_fill

                if isinstance(cell.value, (int, float)) and cell.value >= 1000:
                    cell.number_format = '#,##0'

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                char_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
                if char_len > max_len:
                    max_len = char_len
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output
