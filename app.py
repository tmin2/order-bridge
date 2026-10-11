import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import io
import json
import os
from datetime import datetime
from typing import Any, Optional, Union, Dict, List

# =========================================================================
# 0. Streamlit 설정 및 노션(Notion) 스타일 디자인 시스템
# =========================================================================
st.set_page_config(
    page_title="공동구매 발주서 브릿지 Pro",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');

html, body, [class*="css"] {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", sans-serif !important;
    color: #37352f;
}

/* =========================================================================
   Streamlit Cloud 브랜딩(Fork/GitHub) 및 우측 하단 배지/관리자 버튼 숨김 ⭐
   ========================================================================= */
#MainMenu {visibility: hidden !important; display: none !important;}
footer {visibility: hidden !important; display: none !important;}
[data-testid="stStatusWidget"] {visibility: hidden !important; display: none !important;}
[data-testid="manage-app-button"] {visibility: hidden !important; display: none !important;}
div[data-testid="manage-app-button"] {display: none !important;}
div[class*="viewerBadge"] {display: none !important;}
a[class*="viewerBadge"] {display: none !important;}
div[class*="Profile"] {display: none !important;}
div[class*="FloatingActionButton"] {display: none !important;}
div[class*="bottom-right"] {display: none !important;}
div[class*="BottomRight"] {display: none !important;}
[data-testid="stToolbar"] {visibility: hidden !important; display: none !important;}
header [data-testid="stToolbar"] {display: none !important;}
header a[href*="github.com"] {display: none !important;}
header button[title*="Fork"] {display: none !important;}

/* 모바일 사이드바 열기 버튼(>>)은 안전하게 보존 */
[data-testid="collapsedControl"] {display: block !important; visibility: visible !important;}
[data-testid="stSidebarCollapseButton"] {display: block !important; visibility: visible !important;}

/* Notion Page Title */
.notion-header {
    padding: 12px 0 18px 0;
    margin-bottom: 20px;
    border-bottom: 1px solid #ECECEB;
}
.notion-title {
    font-size: 26px;
    font-weight: 700;
    color: #2F3437;
    display: flex;
    align-items: center;
    gap: 10px;
    letter-spacing: -0.5px;
}
.notion-subtitle {
    font-size: 13.5px;
    color: #787774;
    margin-top: 4px;
}

/* Notion Callout Box */
.notion-callout {
    display: flex;
    align-items: flex-start;
    gap: 12px;
    border-radius: 6px;
    padding: 14px 18px;
    margin: 12px 0 16px 0;
    font-size: 13.5px;
    line-height: 1.55;
    border: 1px solid #E9E9E7;
}
.notion-callout-gray { background-color: #F7F6F3; border-color: #E9E9E7; }
.notion-callout-blue { background-color: #F0F7FD; border-color: #CCE3F8; color: #1D5B90; }
.notion-callout-yellow { background-color: #FDF9EE; border-color: #F7E7B5; color: #7A5805; }
.notion-callout-green { background-color: #F2F8F4; border-color: #CCE7D5; color: #206E3E; }

/* Notion Sidebar Stepper */
.notion-stepper-container {
    background: #FFFFFF;
    border: 1px solid #EAEAE8;
    border-radius: 8px;
    padding: 14px 16px;
    margin-bottom: 16px;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
}
.notion-stepper-title {
    font-size: 13px;
    font-weight: 700;
    color: #787774;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 12px;
    display: flex;
    align-items: center;
    gap: 6px;
}
.notion-step-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 7px 0;
    border-bottom: 1px solid #F4F4F2;
    font-size: 13px;
}
.notion-step-item:last-child {
    border-bottom: none;
}
.notion-step-label {
    color: #37352f;
    font-weight: 500;
}
.notion-badge {
    font-size: 11px;
    font-weight: 600;
    padding: 2px 7px;
    border-radius: 4px;
    display: inline-block;
}
.notion-badge-done { background: #EBF7EE; color: #1E7E34; }
.notion-badge-working { background: #FEF7E6; color: #B25E00; }
.notion-badge-waiting { background: #F1F1EF; color: #787774; }
.notion-badge-skip { background: #F3EEF9; color: #6941C6; }

/* Notion Card */
.notion-card {
    background: #FFFFFF;
    border: 1px solid #E9E9E7;
    border-radius: 8px;
    padding: 18px 22px;
    margin-bottom: 16px;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02);
}

/* 본문 단계 탭(라디오) 대폭 확대 및 노션 탭 스타일 */
div[data-testid="stRadio"] > div[role="radiogroup"] {
    display: flex !important;
    flex-wrap: wrap !important;
    gap: 8px !important;
    background: #F7F6F3 !important;
    padding: 6px 8px !important;
    border-radius: 8px !important;
    border: 1px solid #EAEAE8 !important;
    margin-bottom: 22px !important;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label {
    background: #FFFFFF !important;
    border: 1px solid #E2E2DF !important;
    border-radius: 6px !important;
    padding: 10px 18px !important;
    margin: 0 !important;
    cursor: pointer !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    transition: all 0.15s ease-in-out !important;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label > div[data-testid="stMarkdownContainer"] > p {
    font-size: 19px !important;
    font-weight: 800 !important;
    color: #494541 !important;
    line-height: 1.4 !important;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label:hover {
    background: #F0EFEA !important;
    border-color: #D3D1CB !important;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) {
    background: #2F3437 !important;
    border-color: #2F3437 !important;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) > div[data-testid="stMarkdownContainer"] > p {
    color: #FFFFFF !important;
}
/* 라디오 기본 원형 숨기기 */
div[data-testid="stRadio"] > div[role="radiogroup"] > label > div:first-child {
    display: none !important;
}

/* 본문 단계 헤더 글씨 크기 대폭 확대 (사용자 요청: 단계부분 글씨 크게 ⭐) */
.notion-step-banner {
    background: #F7F6F3;
    border-left: 6px solid #2F3437;
    padding: 18px 24px;
    border-radius: 0 8px 8px 0;
    font-size: 28px !important;
    font-weight: 900 !important;
    color: #2F3437 !important;
    letter-spacing: -0.5px !important;
    margin-bottom: 14px !important;
}
.notion-step-title-big {
    font-size: 28px !important;
    font-weight: 900 !important;
    color: #2F3437 !important;
    letter-spacing: -0.5px !important;
    margin-bottom: 6px !important;
}
.notion-step-desc-big {
    font-size: 16px !important;
    color: #55534E !important;
    line-height: 1.6 !important;
    margin-bottom: 24px !important;
}

/* =========================================================================
   버튼 색상 테마: 다음 단계 이동 (왼쪽 사이드바 활성단계와 100% 동일한 코랄레드 #FF4B4B)
                  템플릿 저장 (동일한 웜레드 톤의 딥 브릭레드 #B83232)
   ========================================================================= */

/* 1. 다음 단계 이동 버튼 (왼쪽 사이드바 활성 버튼과 100% 동일한 비비드 코랄레드 #FF4B4B ⭐) */
div[class*="st-key-btn_goto_"] button,
div[class*="st-key-goto_"] button,
button.agy-btn-coral-next {
    background: #FF4B4B !important;
    background-color: #FF4B4B !important;
    color: #FFFFFF !important;
    border: 1px solid #FF4B4B !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    border-radius: 6px !important;
    box-shadow: 0 2px 6px rgba(255, 75, 75, 0.3) !important;
    transition: all 0.2s ease !important;
}
div[class*="st-key-btn_goto_"] button:hover,
div[class*="st-key-goto_"] button:hover,
button.agy-btn-coral-next:hover {
    background: #FF3333 !important;
    background-color: #FF3333 !important;
    border-color: #E62E2E !important;
    color: #FFFFFF !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(255, 75, 75, 0.45) !important;
}
div[class*="st-key-btn_goto_"] button:active,
div[class*="st-key-goto_"] button:active,
button.agy-btn-coral-next:active {
    transform: translateY(0) !important;
}
div[class*="st-key-btn_goto_"] button p,
div[class*="st-key-goto_"] button p,
button.agy-btn-coral-next p {
    color: #FFFFFF !important;
    font-weight: 700 !important;
}

/* 2. 템플릿 저장 버튼 (다음 단계와 유사한 톤의 차분한 딥 브릭레드 #B83232 ⭐) */
div[class*="st-key-save_"] button,
div[class*="st-key-btn_save_"] button,
div.save-template-btn button,
button.agy-btn-brick-save {
    background: #B83232 !important;
    background-color: #B83232 !important;
    color: #FFFFFF !important;
    border: 1px solid #9E2626 !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    border-radius: 6px !important;
    box-shadow: 0 2px 6px rgba(184, 50, 50, 0.28) !important;
    transition: all 0.2s ease !important;
}
div[class*="st-key-save_"] button:hover,
div[class*="st-key-btn_save_"] button:hover,
div.save-template-btn button:hover,
button.agy-btn-brick-save:hover {
    background: #A32020 !important;
    background-color: #A32020 !important;
    border-color: #8A1515 !important;
    color: #FFFFFF !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(184, 50, 50, 0.38) !important;
}
div[class*="st-key-save_"] button:active,
div[class*="st-key-btn_save_"] button:active,
button.agy-btn-brick-save:active {
    transform: translateY(0) !important;
}
div[class*="st-key-save_"] button p,
div[class*="st-key-btn_save_"] button p,
div.save-template-btn button p,
button.agy-btn-brick-save p {
    color: #FFFFFF !important;
    font-weight: 700 !important;
}

/* 3. 최종 발주서 엑셀 다운로드 버튼 (폰트 확대 & 다홍색 #FF4B4B 깜빡임/펄스 애니메이션 ⭐) */
@keyframes coral-blink-pulse {
    0%, 100% {
        background-color: #FF4B4B !important;
        border-color: #E63939 !important;
        box-shadow: 0 3px 12px rgba(255, 75, 75, 0.4) !important;
        opacity: 0.95;
    }
    50% {
        background-color: #FF1E1E !important;
        border-color: #C71414 !important;
        box-shadow: 0 0 24px 8px rgba(255, 75, 75, 0.75) !important;
        opacity: 1;
        transform: scale(1.02);
    }
}

div.final-download-btn-pulse button,
div.final-download-btn-pulse div[data-testid="stDownloadButton"] button,
div[class*="st-key-btn_download_final_"] button {
    background: #FF4B4B !important;
    background-color: #FF4B4B !important;
    color: #FFFFFF !important;
    border: 2px solid #E63939 !important;
    border-radius: 8px !important;
    font-size: 19px !important;
    font-weight: 800 !important;
    letter-spacing: -0.3px !important;
    height: 56px !important;
    min-height: 56px !important;
    animation: coral-blink-pulse 1.8s infinite ease-in-out !important;
    transition: all 0.2s ease !important;
    box-shadow: 0 3px 12px rgba(255, 75, 75, 0.4) !important;
}

div.final-download-btn-pulse button:hover,
div.final-download-btn-pulse div[data-testid="stDownloadButton"] button:hover,
div[class*="st-key-btn_download_final_"] button:hover {
    background: #E62020 !important;
    background-color: #E62020 !important;
    border-color: #BA1212 !important;
    color: #FFFFFF !important;
    transform: scale(1.03) !important;
    box-shadow: 0 6px 26px rgba(255, 75, 75, 0.85) !important;
}

div.final-download-btn-pulse button p,
div.final-download-btn-pulse button span,
div[class*="st-key-btn_download_final_"] button p,
div[class*="st-key-btn_download_final_"] button span {
    color: #FFFFFF !important;
    font-size: 19px !important;
    font-weight: 800 !important;
    letter-spacing: -0.3px !important;
    margin: 0 !important;
    padding: 0 !important;
}

/* 모든 액션 버튼 높이 및 행 정렬 일치 ⭐ */
div[class*="st-key-btn_"] button,
div[class*="st-key-save_"] button {
    height: 44px !important;
    min-height: 44px !important;
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    margin-top: 0 !important;
    margin-bottom: 0 !important;
    vertical-align: middle !important;
}

div[class*="st-key-btn_"] button p,
div[class*="st-key-save_"] button p {
    margin: 0 !important;
    padding: 0 !important;
    line-height: normal !important;
}

/* 사이드바 진행 단계 버튼: 2줄 표시 및 자동 높이 보장 ⭐ */
section[data-testid="stSidebar"] div[class*="st-key-sidebar_step_nav_btn_"] button {
    height: auto !important;
    min-height: 52px !important;
    padding: 8px 12px !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: flex-start !important;
    justify-content: center !important;
    text-align: left !important;
    margin-bottom: 4px !important;
}

section[data-testid="stSidebar"] div[class*="st-key-sidebar_step_nav_btn_"] button div[data-testid="stMarkdownContainer"] {
    width: 100% !important;
    text-align: left !important;
}

section[data-testid="stSidebar"] div[class*="st-key-sidebar_step_nav_btn_"] button p {
    white-space: pre-line !important;
    word-break: break-all !important;
    text-align: left !important;
    margin: 0 !important;
    line-height: 1.38 !important;
    font-size: 13px !important;
}

/* 사이드바 백업 복원 업로더 가운데 정렬 ⭐ */
section[data-testid="stSidebar"] div[data-testid="stFileUploader"] {
    text-align: center !important;
}
section[data-testid="stSidebar"] div[data-testid="stFileUploader"] label {
    text-align: center !important;
    display: block !important;
    width: 100% !important;
    margin-bottom: 6px !important;
}
section[data-testid="stSidebar"] div[data-testid="stFileUploaderDropzone"] {
    text-align: center !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    padding: 14px 10px !important;
}
section[data-testid="stSidebar"] div[data-testid="stFileUploaderDropzone"] div {
    text-align: center !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
}
section[data-testid="stSidebar"] div[data-testid="stFileUploaderDropzone"] button {
    margin: 0 auto !important;
    display: inline-block !important;
}
section[data-testid="stSidebar"] div[data-testid="stFileUploaderDropzone"] small {
    text-align: center !important;
    display: block !important;
    width: 100% !important;
}
</style>

<script>
function applyAgyButtonStyles() {
    try {
        const buttons = document.querySelectorAll('button');
        buttons.forEach(btn => {
            if (btn.closest('section[data-testid="stSidebar"]')) return;
            const txt = (btn.innerText || btn.textContent || '').trim();
            if (txt.includes('템플릿 저장') || txt.includes('영구 저장') || txt.includes('템플릿 등록')) {
                btn.classList.add('agy-btn-brick-save');
                btn.classList.remove('agy-btn-coral-next');
            } else if ((txt.includes('단계') && (txt.includes('이동') || txt.includes('생성'))) || txt.includes('5단계 최종 발주서')) {
                btn.classList.add('agy-btn-coral-next');
                btn.classList.remove('agy-btn-brick-save');
            }
        });
    } catch (e) {}
}
applyAgyButtonStyles();
if (!window._agyObserverAttached) {
    window._agyObserverAttached = true;
    const obs = new MutationObserver(applyAgyButtonStyles);
    obs.observe(document.body, { childList: true, subtree: true });
}
</script>
""", unsafe_allow_html=True)
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
BRAND_FILE = os.path.join(CURRENT_DIR, "brand_templates.json")
SELLER_MAP_FILE = os.path.join(CURRENT_DIR, "seller_mappings.json")
EVENT_TEMPLATE_FILE = os.path.join(CURRENT_DIR, "event_templates.json")
BRAND_PRODUCT_FILE = os.path.join(CURRENT_DIR, "brand_product_master.json")

STANDARD_FIELDS = {
    "order_time": "주문일시 / 결제일시",
    "order_id": "주문번호 / 관리코드",
    "buyer_name": "주문자(구매자)명",
    "buyer_phone": "주문자 연락처",
    "receiver_name": "수령인(받는분)명",
    "receiver_phone": "수령인 연락처",
    "zipcode": "우편번호",
    "address": "기본 배송주소",
    "address_detail": "상세 배송주소",
    "product_name": "상품명 / 품목명 (실제 한글 상품명)",
    "option_name": "옵션명 / 선택사항",
    "quantity": "수량",
    "price": "결제금액 / 판매가 (없으면 본사 단가 자동적용)",
    "shipping_memo": "배송메모 / 요청사항",
    "seller_channel": "판매처 / 쇼핑몰명 (엑셀 열 매칭)"
}

STANDARD_COLUMN_ALIASES = {
    "order_time": ["주문일시", "결제일시", "주문일자", "결제일", "주문시간", "order_date", "date", "주문일"],
    "order_id": ["쇼핑몰 주문번호", "주문번호", "관리번호", "주문id", "order_id", "주문코드", "쇼핑몰주문번호"],
    "buyer_name": ["주문자명", "구매자", "주문자", "고객명", "구매자명", "주문자 이름"],
    "buyer_phone": ["주문자연락처", "구매자전화번호", "주문자핸드폰", "구매자연락처", "주문자 전화번호", "주문자 휴대폰"],
    "receiver_name": ["수취인명", "수령인", "받는분", "받으시는분", "수령인명", "받는사람", "수령자", "수하인", "받는분 성명"],
    "receiver_phone": ["수취인전화번호1", "수취인연락처", "수령인연락처", "받는분연락처", "받으시는분전화번호", "전화번호", "휴대폰", "연락처", "핸드폰", "수취인핸드폰", "수령자연락처", "수령자전화번호", "수령자핸드폰", "휴대폰번호", "전화번호1", "휴대폰1", "받는분 전화번호"],
    "zipcode": ["우편번호", "우편", "zipcode", "postcode"],
    "address": ["수취인주소", "기본주소", "배송주소", "배송지", "주소", "배송지주소", "수령인주소", "수령자주소", "수령자 주소", "배송지 기본주소", "기본배송주소", "배송주소지", "주소1"],
    "address_detail": ["상세주소", "배송지상세", "상세배송주소", "배송지상세주소", "배송지 상세주소", "주소2"],
    "product_name": ["품목명", "주문상품", "상품약어", "상품명", "품목", "상품", "product_name", "주문품목"],
    "option_name": ["옵션", "선택옵션", "옵션정보", "옵션명", "규격", "옵션사항", "선택사항", "옵션코드 기입", "옵션항목"],
    "quantity": ["수량", "주문수량", "수량(개)", "qty", "개수", "구매수량", "주문개수"],
    "price": ["결제금액", "판매가", "결제액", "단가", "주문금액", "금액", "구매금액"],
    "shipping_memo": ["배송메세지", "배송메모", "배송요청사항", "배송 요청사항", "배송시 요청사항", "메모", "배송요청", "요청사항"],
    "seller_channel": ["판매처", "쇼핑몰", "쇼핑몰명", "판매처명", "셀러명", "채널", "채널명", "판매채널", "몰명", "store", "channel"]
}

def extract_date_from_filename(filename):
    """파일명에서 발주일자(YYYY-MM-DD, MMDD, M월 D일 등) 자동 추출"""
    if not filename:
        return datetime.now().date()
    curr_year = datetime.now().year
    
    # 1. YYYY-MM-DD, YYYY.MM.DD, YYYY_MM_DD
    m1 = re.search(r'(20\d{2})[._-](\d{1,2})[._-](\d{1,2})', filename)
    if m1:
        try:
            return datetime(int(m1.group(1)), int(m1.group(2)), int(m1.group(3))).date()
        except ValueError:
            pass
            
    # 2. YYYYMMDD (8자리 숫자)
    m2 = re.search(r'(20\d{2})(\d{2})(\d{2})', filename)
    if m2:
        try:
            return datetime(int(m2.group(1)), int(m2.group(2)), int(m2.group(3))).date()
        except ValueError:
            pass

    # 3. M월 D일
    m3 = re.search(r'(\d{1,2})월\s*(\d{1,2})일', filename)
    if m3:
        try:
            return datetime(curr_year, int(m3.group(1)), int(m3.group(2))).date()
        except ValueError:
            pass

    # 4. MMDD 4자리 숫자 (예: 1008 in '리오마레_발주서 1008 송장')
    m4 = re.findall(r'(?<!\d)(\d{2})(\d{2})(?!\d)', filename)
    for mm_str, dd_str in m4:
        mm, dd = int(mm_str), int(dd_str)
        if 1 <= mm <= 12 and 1 <= dd <= 31:
            try:
                return datetime(curr_year, mm, dd).date()
            except ValueError:
                pass
                
    # 5. MM.DD or MM-DD or MM_DD
    m5 = re.search(r'(?<!\d)(\d{1,2})[._-](\d{1,2})(?!\d)', filename)
    if m5:
        mm, dd = int(m5.group(1)), int(m5.group(2))
        if 1 <= mm <= 12 and 1 <= dd <= 31:
            try:
                return datetime(curr_year, mm, dd).date()
            except ValueError:
                pass

    return datetime.now().date()

def extract_clean_seller_from_filename(filename):
    """파일명에서 날짜/발주서/송장 등 수식어를 제거하고 순수 셀러/판매처명 추출"""
    if not filename:
        return ""
    base = re.sub(r'\.[a-zA-Z0-9]+$', '', filename)
    base = re.sub(r'(20\d{2})[._-]?\d{1,2}[._-]?\d{1,2}', '', base)
    base = re.sub(r'\b\d{4}\b', '', base)
    base = re.sub(r'\d{1,2}월\s*\d{1,2}일', '', base)
    base = re.sub(r'[_\s-]*(발주서|발주|주문서|주문|송장|배송|양식|출고|리스트|sheet|order).*$', '', base, flags=re.IGNORECASE)
    base = re.sub(r'^[_\s-]+|[_\s-]+$', '', base)
    return base.strip()

# =========================================================================
# 1. 템플릿 저장 & 로드 함수
# =========================================================================
def load_brand_templates():
    if os.path.exists(BRAND_FILE):
        try:
            with open(BRAND_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k, v in data.items():
                    if isinstance(v, dict):
                        if "template_name" not in v:
                            v["template_name"] = k
                        if "brand_name" not in v:
                            v["brand_name"] = k
                return data
        except Exception:
            pass
    return {
        "리오마레": {
            "template_name": "리오마레",
            "brand_name": "리오마레",
            "columns": ["날짜", "판매처", "주문번호", "수령자", "수령자주소", "수령자전화번호", "수령자핸드폰", "상품명", "수량", "송장번호", "배송메세지"],
            "gift_rule": "add_row",
            "shipping_config": {
                "bundle_targets": ["수취인명", "주소지", "연락처"],
                "free_shipping_threshold": 50000,
                "base_shipping_fee": 3000,
                "jeju_extra_fee": 3000
            }
        },
        "호무로": {
            "template_name": "호무로",
            "brand_name": "호무로",
            "columns": ["쇼핑몰명", "쇼핑몰 주문번호", "사방넷 주문번호", "옵션코드 기입", "사방넷상품코드", "상품약어", "수량", "수취인명", "수취인전화번호1", "수취인주소", "배송메세지", "물류처명", "송장번호"],
            "gift_rule": "add_row",
            "shipping_config": {
                "bundle_targets": ["수취인명", "주소지", "연락처"],
                "free_shipping_threshold": 50000,
                "base_shipping_fee": 3000,
                "jeju_extra_fee": 3000
            }
        },
        "기본_표준양식": {
            "template_name": "기본_표준양식",
            "brand_name": "기본_표준양식",
            "columns": ["순번", "셀러명", "주문일시", "주문번호", "수령인", "연락처", "우편번호", "주소", "브랜드상품명", "옵션명", "수량", "제주도서여부", "이벤트색인", "배송메모"],
            "gift_rule": "add_row",
            "shipping_config": {
                "bundle_targets": ["수취인명", "주소지", "연락처"],
                "free_shipping_threshold": 50000,
                "base_shipping_fee": 3000,
                "jeju_extra_fee": 3000
            }
        }
    }

def save_brand_templates(data):
    with open(BRAND_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def save_brand_template(tmpl_key, columns, brand_name="", gift_rule="add_row", shipping_config=None):
    data = load_brand_templates()
    clean_tmpl_key = str(tmpl_key).strip()
    clean_brand = str(brand_name).strip() if brand_name else clean_tmpl_key
    entry = {
        "template_name": clean_tmpl_key,
        "brand_name": clean_brand,
        "columns": columns,
        "gift_rule": gift_rule
    }
    if shipping_config is not None:
        entry["shipping_config"] = shipping_config
    elif clean_tmpl_key in data and "shipping_config" in data[clean_tmpl_key]:
        entry["shipping_config"] = data[clean_tmpl_key]["shipping_config"]
    else:
        entry["shipping_config"] = {
            "bundle_targets": ["수취인명", "주소지", "연락처"],
            "free_shipping_threshold": 50000,
            "base_shipping_fee": 3000,
            "jeju_extra_fee": 3000
        }
    data[clean_tmpl_key] = entry
    save_brand_templates(data)

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

def extract_leading_number(text):
    """
    텍스트 앞부분의 번호(1., 2-, (1), [1], 1 등)를 숫자로 추출.
    '필수옵션=2. 롤티슈 2팩' 같은 접두사도 깔끔하게 제거하고 추출.
    """
    if not text:
        return None
    s = str(text).strip()
    s = re.sub(r'^(?:필수옵션\s*=\s*|선택옵션\s*=\s*|옵션\s*:\s*|옵션\s*=\s*)', '', s, flags=re.IGNORECASE).strip()
    m = re.match(r'^(?:\[|\()?(\d+)(?:\.|\)|\]|\-|\:|\s|번|$)', s)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            pass
    m2 = re.match(r'^(\d+)', s)
    if m2:
        try:
            return int(m2.group(1))
        except ValueError:
            pass
    return None

def save_seller_mapping(tmpl_name, seller_name, brand_name, column_mapping, product_mappings=None, custom_seller_name="", custom_order_date="", force_custom_date=True, product_order=None):
    data = load_seller_mappings()
    existing = data.get(tmpl_name, {})
    if product_mappings is None:
        product_mappings = existing.get("product_mappings", {})

    final_custom_seller = custom_seller_name if custom_seller_name else existing.get("custom_seller_name", "")
    final_custom_date = custom_order_date if custom_order_date else existing.get("custom_order_date", "")
    final_force_date = force_custom_date if force_custom_date is not None else existing.get("force_custom_date", True)
    final_product_order = product_order if product_order is not None else existing.get("product_order", [])

    entry = {
        "template_name": tmpl_name,
        "seller_name": seller_name or existing.get("seller_name", tmpl_name),
        "brand_name": brand_name or existing.get("brand_name", ""),
        "column_mapping": column_mapping if column_mapping is not None else existing.get("column_mapping", {}),
        "product_mappings": product_mappings,
        "product_order": final_product_order,
        "custom_seller_name": final_custom_seller,
        "custom_order_date": final_custom_date,
        "force_custom_date": final_force_date,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    if "event_config" in existing:
        entry["event_config"] = existing["event_config"]

    data[tmpl_name] = entry
    save_seller_mappings(data)


def load_brand_product_master():
    if os.path.exists(BRAND_PRODUCT_FILE):
        try:
            with open(BRAND_PRODUCT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_brand_product_master(data):
    with open(BRAND_PRODUCT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def update_brand_product_master_from_pm(brand_name, product_mapping_dict):
    """3단계에서 매핑/저장된 품목들을 해당 브랜드의 공식 품목 마스터 DB에 누적 저장/업데이트"""
    if not brand_name or not product_mapping_dict:
        return
    master_data = load_brand_product_master()
    clean_b = str(brand_name).strip()
    if clean_b not in master_data:
        master_data[clean_b] = {}
    
    brand_dict = master_data[clean_b]
    for k_pair, info in product_mapping_dict.items():
        bp = str(info.get("brand_product", "")).strip()
        bo = str(info.get("brand_option", "")).strip()
        raw_p_part = k_pair.split("||")[0].strip()
        # 셀러 원본 상품명이 그대로 들어가 있거나 셀러명(에이핏 등)이 포함된 경우 마스터에 저장 금지
        if not bp or bp == raw_p_part or "에이핏" in bp:
            continue
        try:
            up = int(float(str(info.get("unit_price", 0)).replace(",", "").replace("원", "")))
        except Exception:
            up = 0
        code = str(info.get("code", "")).strip()
        
        is_pre = bool(info.get("is_pre_registered", False) or str(k_pair).startswith("__official__"))
        is_chk = bool(info.get("is_matched", False) or up > 0 or is_pre)
        if bp and (up > 0 or is_chk or is_pre):
            item_entry = {
                "brand_product": bp,
                "brand_option": bo,
                "unit_price": up,
                "code": code,
                "is_matched": bool(info.get("is_matched", False)),
                "is_pre_registered": is_pre,
                "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M")
            }
            clean_k = f"{bp}||{bo}"
            brand_dict[clean_k] = item_entry
            if not str(k_pair).startswith("__official__"):
                brand_dict[k_pair] = item_entry

    master_data[clean_b] = brand_dict
    save_brand_product_master(master_data)

def get_all_known_brand_mappings(brand_name):
    """
    해당 브랜드(active_brand)에 대해 저장된 모든 공식 상품명, 옵션, 단가, 코드 정보를 종합 수집:
    1) brand_product_master.json 에 저장된 데이터
    2) seller_mappings.json 에서 동일 브랜드(brand_name)로 매핑된 모든 템플릿의 product_mappings
    """
    known_pool = {}
    target_b = str(brand_name).strip()
    
    # 1. seller_mappings.json 에서 수집
    all_sm = load_seller_mappings()
    for s_key, s_data in all_sm.items():
        s_brand = str(s_data.get("brand_name", "")).strip()
        if s_brand and (s_brand == target_b or target_b in s_brand or s_brand in target_b):
            pm = s_data.get("product_mappings", {})
            for k_pair, info in pm.items():
                if info.get("brand_product") and (int(info.get("unit_price", 0)) > 0 or info.get("is_matched") or info.get("is_pre_registered")):
                    known_pool[k_pair] = info.copy()

    # 2. brand_product_master.json 에서 수집 (우선순위 높음)
    master_data = load_brand_product_master()
    for b_key, b_dict in master_data.items():
        b_clean = str(b_key).strip()
        if b_clean and (b_clean == target_b or target_b in b_clean or b_clean in target_b):
            for k_pair, info in b_dict.items():
                if info.get("brand_product") and (int(info.get("unit_price", 0)) > 0 or info.get("is_matched") or info.get("is_pre_registered")):
                    known_pool[k_pair] = info.copy()

    return known_pool

def auto_match_products_against_brand_pool(unique_items, brand_pool):
    """
    unique_items [(raw_p, raw_o), ...] 와 brand_pool 간의 스마트 자동 대조 및 매칭.
    반환: (matched_results, matched_count)
    """
    matched_results = {}
    matched_count = 0

    if not brand_pool or not unique_items:
        return matched_results, 0

    def clean_text(s):
        s = re.sub(r'^(에이핏\s*[xX:\-]\s*|에이핏\s*x\s*리오마레\s*[:\-]?\s*|에이핏\s*|구성품\s*[:：]?|단품\s*[:：]?|세트\s*[:：]?|\[.*?\]|\(.*?\))\s*', '', str(s), flags=re.IGNORECASE)
        return re.sub(r'[^0-9a-zA-Z가-힣]', '', s).lower()

    pool_entries = []
    for k_pair, info in brand_pool.items():
        parts = k_pair.split("||")
        p_raw = parts[0] if len(parts) > 0 else ""
        o_raw = parts[1] if len(parts) > 1 else ""
        bp = str(info.get("brand_product", "")).strip()
        bo = str(info.get("brand_option", "")).strip()
        # bp가 비어있거나 에이핏이 들어간 것은 제외
        if not bp or "에이핏" in bp:
            continue
        up = int(info.get("unit_price", 0))
        code = str(info.get("code", "")).strip()
        pool_entries.append({
            "p_raw": p_raw,
            "o_raw": o_raw,
            "bp": bp,
            "bo": bo,
            "up": up,
            "code": code,
            "aliases": [p_raw, bp] if p_raw != bp else [bp]
        })

    for raw_p, raw_o in unique_items:
        k_pair = f"{raw_p}||{raw_o}"
        
        # 1. 완전 일치
        if k_pair in brand_pool:
            inf = brand_pool[k_pair]
            bp_val = str(inf.get("brand_product", "")).strip()
            # 셀러 원본 상품명이 아니고 순수한 공식 브랜드 상품명일 때만 채움
            if bp_val and "에이핏" not in bp_val:
                matched_results[k_pair] = {
                    "brand_product": bp_val,
                    "brand_option": str(inf.get("brand_option", raw_o)).strip(),
                    "unit_price": int(inf.get("unit_price", 0)),
                    "is_matched": True
                }
                matched_count += 1
                continue

        c_raw = clean_text(raw_p)
        clean_ro = re.sub(r'[^0-9a-zA-Z]', '', str(raw_o)).lower()
        raw_specs = set(re.findall(r'\d+(?:g|ml|kg|캔|개|입|세트|box)?', str(raw_p).lower()))
        
        best_entry = None
        best_score = 0

        for entry in pool_entries:
            score = 0
            # 코드 대조
            e_code = re.sub(r'[^0-9a-zA-Z]', '', str(entry.get("code", ""))).lower()
            if e_code and len(e_code) >= 4:
                if e_code in clean_ro or e_code in clean_text(raw_p):
                    score += 2000

            for alias in entry["aliases"]:
                if not alias: continue
                c_tgt = clean_text(alias)
                if c_raw == c_tgt:
                    score = max(score, 1000)
                elif c_tgt in c_raw or c_raw in c_tgt:
                    score = max(score, 500)

                t_specs = set(re.findall(r'\d+(?:g|ml|kg|캔|개|입|세트|box)?', alias.lower()))
                common_specs = raw_specs.intersection(t_specs)
                if common_specs:
                    score += len(common_specs) * 80

                for s1 in raw_specs:
                    for s2 in t_specs:
                        if s1 != s2:
                            u1 = re.sub(r'\d', '', s1)
                            u2 = re.sub(r'\d', '', s2)
                            if u1 and u1 == u2:
                                score -= 120

                r_words = set(re.findall(r'[가-힣]{2,}', str(raw_p)))
                t_words = set(re.findall(r'[가-힣]{2,}', alias))
                common_words = r_words.intersection(t_words)
                score += len(common_words) * 60

            if score > best_score:
                best_score = score
                best_entry = entry

        if best_entry is not None and best_score >= 200:
            matched_results[k_pair] = {
                "brand_product": best_entry["bp"],
                "brand_option": best_entry["bo"] if best_entry["bo"] else raw_o,
                "unit_price": best_entry["up"],
                "is_matched": True
            }
            matched_count += 1

    return matched_results, matched_count

def delete_seller_mapping(tmpl_name):
    data = load_seller_mappings()
    if tmpl_name in data:
        del data[tmpl_name]
        save_seller_mappings(data)
        return True
    return False

def get_seller_mapping(tmpl_name):
    data = load_seller_mappings()
    if tmpl_name in data:
        return data[tmpl_name]
    for key, val in data.items():
        if key in tmpl_name or tmpl_name in key:
            return val
    return None

def find_best_seller_mapping(file_name, brand_name=""):
    data = load_seller_mappings()
    if not data or not file_name or not str(file_name).strip():
        return None, None
    clean_f = str(file_name).strip()
    if clean_f in data:
        return clean_f, data[clean_f]
    for k, v in data.items():
        if not k or not str(k).strip() or len(str(k).strip()) < 2:
            continue
        v_brand = v.get("brand_name", "")
        v_seller = v.get("seller_name", k)
        if brand_name and v_brand == brand_name and v_seller and len(v_seller) >= 2:
            if v_seller in clean_f or clean_f in v_seller:
                return k, v
    for k, v in data.items():
        if not k or not str(k).strip() or len(str(k).strip()) < 2:
            continue
        if k in clean_f or clean_f in k:
            return k, v
    return None, None

def load_event_templates():
    if os.path.exists(EVENT_TEMPLATE_FILE):
        try:
            with open(EVENT_TEMPLATE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                # 하위 호환: is_pinned 필드가 없으면 True로 기본 설정 (자주 쓰는 템플릿 기본 노출)
                for k, v in data.items():
                    if "is_pinned" not in v:
                        v["is_pinned"] = True
                return data
        except Exception:
            pass
    return {
        "기본 (이벤트 없음)": {
            "template_name": "기본 (이벤트 없음)",
            "config": {
                "color_event": "#FFF2CC",
                "color_gift": "#E2EFDA",
                "first_come_first_served": {"enabled": False},
                "min_amount": {"enabled": False},
                "special_product": {"enabled": False},
                "custom_events": []
            },
            "is_pinned": True,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M")
        }
    }

def save_event_templates(data):
    with open(EVENT_TEMPLATE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def save_event_template(tmpl_name, config, is_pinned=None):
    data = load_event_templates()
    existing_entry = data.get(tmpl_name, {})
    existing_pin = existing_entry.get("is_pinned", True)
    data[tmpl_name] = {
        "template_name": tmpl_name,
        "config": config,
        "is_pinned": is_pinned if is_pinned is not None else existing_pin,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    save_event_templates(data)

def toggle_event_template_pinned(tmpl_name, is_pinned):
    data = load_event_templates()
    if tmpl_name in data:
        data[tmpl_name]["is_pinned"] = is_pinned
        save_event_templates(data)
        return True
    return False

def delete_event_template(tmpl_name):
    data = load_event_templates()
    if tmpl_name in data:
        del data[tmpl_name]
        save_event_templates(data)
        return True
    return False

def get_event_template(tmpl_name):
    data = load_event_templates()
    return data.get(tmpl_name)

def parse_header_text(raw_text):
    if not raw_text:
        return []
    cleaned = raw_text.replace("\t", ",").replace("\n", ",")
    items = [c.strip() for c in cleaned.split(",") if c.strip()]
    return items

def export_full_backup_json():
    """전체 브랜드 템플릿, 셀러 매핑, 이벤트 템플릿, 공식 품목 마스터를 하나의 JSON으로 백업"""
    backup = {
        "version": "2.1",
        "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "brand_templates": load_brand_templates(),
        "seller_mappings": load_seller_mappings(),
        "event_templates": load_event_templates(),
        "brand_product_master": load_brand_product_master()
    }
    return json.dumps(backup, ensure_ascii=False, indent=2)

def import_full_backup_json(json_str):
    """백업 JSON 파일로부터 브랜드 템플릿, 셀러 매핑, 이벤트 템플릿, 공식 품목 마스터 복원"""
    try:
        data = json.loads(json_str)
        if "brand_templates" in data:
            save_brand_templates(data["brand_templates"])
        if "seller_mappings" in data:
            save_seller_mappings(data["seller_mappings"])
        if "event_templates" in data:
            save_event_templates(data["event_templates"])
        if "brand_product_master" in data:
            save_brand_product_master(data["brand_product_master"])
        return True, "성공적으로 복원되었습니다."
    except Exception as e:
        return False, f"복원 실패: {str(e)}"


# =========================================================================
# 1-1. 일자별 발주 데이터 누적 저장소 (동일 셀러+브랜드 템플릿 연동)
# =========================================================================
ORDER_HISTORY_DIR = os.path.join(CURRENT_DIR, "order_history")

def get_order_history_file(seller_name, brand_name):
    os.makedirs(ORDER_HISTORY_DIR, exist_ok=True)
    clean_s = re.sub(r'[^a-zA-Z0-9가-힣_-]', '_', str(seller_name or "기본셀러").strip())
    clean_b = re.sub(r'[^a-zA-Z0-9가-힣_-]', '_', str(brand_name or "기본브랜드").strip())
    return os.path.join(ORDER_HISTORY_DIR, f"{clean_s}_{clean_b}.json")

get_accumulated_filepath = get_order_history_file

def load_accumulated_orders(seller_name, brand_name):
    fpath = get_order_history_file(seller_name, brand_name)
    if os.path.exists(fpath):
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                data = json.load(f)
                if data:
                    df = pd.DataFrame(data)
                    return df
        except Exception:
            pass
    return pd.DataFrame()

def save_accumulated_orders(seller_name, brand_name, df_new):
    if df_new is None or df_new.empty:
        return
    fpath = get_order_history_file(seller_name, brand_name)
    df_old = load_accumulated_orders(seller_name, brand_name)
    
    df_save_new = df_new.copy()
    if "order_time_obj" in df_save_new.columns:
        df_save_new["order_time_obj"] = df_save_new["order_time_obj"].astype(str)
        
    if not df_old.empty:
        if "order_time_obj" in df_old.columns:
            df_old["order_time_obj"] = df_old["order_time_obj"].astype(str)
        merged = pd.concat([df_old, df_save_new], ignore_index=True)
        subset_cols = [c for c in ["order_id", "product_name", "option_name", "order_date"] if c in merged.columns]
        merged.drop_duplicates(subset=subset_cols, keep="last", inplace=True)
    else:
        merged = df_save_new

    records = merged.to_dict(orient="records")
    with open(fpath, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2, default=str)

def clear_accumulated_orders(seller_name, brand_name):
    fpath = get_order_history_file(seller_name, brand_name)
    if os.path.exists(fpath):
        try:
            os.remove(fpath)
            return True
        except Exception:
            pass
    return False

def create_empty_brand_template_excel(columns, brand_name="브랜드사양식"):
    """브랜드사의 헤더 컬럼만 포함된 빈 양식 엑셀 파일 생성"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{brand_name}_발주양식"
    
    header_fill = PatternFill(start_color="1E3D59", end_color="1E3D59", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    
    for c_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = max(len(str(col_name)) * 2 + 5, 12)
        
    ws.row_dimensions[1].height = 28
    ws.views.sheetView[0].showGridLines = True
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

def export_product_mapping_excel(product_mappings, brand_master_dict=None):
    """현재 매핑된 품목 & 본사 판매단가 마스터(사전등록 공식목록 포함)를 엑셀 파일로 내보내기"""
    rows = []
    seen_official = set()
    for key, val in product_mappings.items():
        is_pre = bool(val.get("is_pre_registered") or str(key).startswith("__official__"))
        clean_key = str(key).replace("__official__", "")
        parts = clean_key.split("||")
        raw_p = parts[0] if len(parts) > 0 else ""
        raw_o = parts[1] if len(parts) > 1 else ""
        bp = str(val.get("brand_product", "")).strip() or raw_p
        bo = str(val.get("brand_option", "")).strip() or raw_o
        up = int(val.get("unit_price", 0))
        code = str(val.get("code", "")).strip()

        if is_pre or "사전등록" in raw_p or "발주서 미포함" in raw_p:
            status_str = "⏳ 매칭대기 (공식목록)"
            disp_raw_p = "(사전등록 공식품목)"
            disp_raw_o = "(발주서 미포함)"
        elif val.get("is_matched") or up > 0:
            status_str = "✅ 매칭완료"
            disp_raw_p = raw_p
            disp_raw_o = raw_o
        else:
            status_str = "⚠️ 미매칭 (단가미입력)"
            disp_raw_p = raw_p
            disp_raw_o = raw_o

        if bp:
            seen_official.add(f"{bp}||{bo}")

        rows.append({
            "브랜드_공식상품명": bp,
            "브랜드_공식옵션": bo,
            "상품/옵션코드": code,
            "본사_판매단가(원)": up,
            "매핑상태": status_str,
            "셀러_원래상품명": disp_raw_p,
            "셀러_원래옵션명": disp_raw_o
        })

    # 누적 마스터 풀(brand_master_dict)에 등록된 공식 품목 중 saved_pm에 아직 없는 품목도 함께 내보내기
    if brand_master_dict:
        for b_key, b_info in brand_master_dict.items():
            bp = str(b_info.get("brand_product", "")).strip()
            bo = str(b_info.get("brand_option", "")).strip()
            up = int(b_info.get("unit_price", 0))
            code = str(b_info.get("code", "")).strip()
            if not bp or "에이핏" in bp:
                continue
            pair_key = f"{bp}||{bo}"
            if pair_key not in seen_official:
                seen_official.add(pair_key)
                rows.append({
                    "브랜드_공식상품명": bp,
                    "브랜드_공식옵션": bo,
                    "상품/옵션코드": code,
                    "본사_판매단가(원)": up,
                    "매핑상태": "⏳ 매칭대기 (공식목록)",
                    "셀러_원래상품명": "(사전등록 공식품목)",
                    "셀러_원래옵션명": "(발주서 미포함)"
                })

    df_exp = pd.DataFrame(rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_exp.to_excel(writer, sheet_name="품목단가_마스터매핑", index=False)
        ws = writer.sheets["품목단가_마스터매핑"]
        
        # openpyxl 서식 및 열 너비 자동 조정
        header_fill = PatternFill(start_color="1E3D59", end_color="1E3D59", fill_type="solid")
        header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
        data_font = Font(name="맑은 고딕", size=10)
        border_thin = Border(
            left=Side(style='thin', color='DDDDDD'),
            right=Side(style='thin', color='DDDDDD'),
            top=Side(style='thin', color='DDDDDD'),
            bottom=Side(style='thin', color='DDDDDD')
        )
        
        for col_idx, col_name in enumerate(df_exp.columns, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # 열 너비
            max_len = max(len(str(col_name)) * 2, 12)
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = max_len + 4
            
        for row_idx in range(2, len(df_exp) + 2):
            for col_idx in range(1, len(df_exp.columns) + 1):
                c = ws.cell(row=row_idx, column=col_idx)
                c.font = data_font
                c.border = border_thin
                col_name = df_exp.columns[col_idx - 1]
                if col_name == "본사_판매단가(원)":
                    c.number_format = '#,##0'
                    c.alignment = Alignment(horizontal="right", vertical="center")
                elif col_name in ["상품/옵션코드", "매핑상태"]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")

        ws.row_dimensions[1].height = 28
        ws.views.sheetView[0].showGridLines = True

    buf.seek(0)
    return buf.getvalue()


# =========================================================================
# 2. 비즈니스 로직 및 데이터 처리 엔진
# =========================================================================
def detect_column_mapping(df_columns):
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
                        if std_key == "product_name" and "코드" in norm_c:
                            continue
                        col_map[std_key] = orig_c
                        found = True
                        break
                if found:
                    break
        if not found:
            col_map[std_key] = "(선택안함)"
    return col_map

def is_jeju_island(address, zipcode=""):
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

def clean_scalar_str(val, default=""):
    """주문번호, 우편번호, 연락처 등 숫자 끝에 .0이 붙는 부동소수점 오염 및 NaN/None을 말끔히 정제하여 순수 문자열로 반환"""
    if pd.isna(val) or val is None:
        return default
    if isinstance(val, (float, np.floating)):
        if np.isnan(val) or np.isinf(val):
            return default
        if val.is_integer():
            return str(int(val))
        return f"{val:.8f}".rstrip("0").rstrip(".")
    s_val = str(val).strip()
    if s_val.lower() in ["", "nan", "none", "null", "<na>"]:
        return default
    if s_val.endswith(".0"):
        stem = s_val[:-2]
        if stem.replace("-", "").isalnum():
            return stem
    return s_val

def sanitize_dataframe_columns(df):
    """모든 컬럼명을 안전하고 고유한 문자열로 정돈하고 숫자 뒤의 .0 부동소수점 오염 제거"""
    if df is None or not isinstance(df, pd.DataFrame):
        return df
    new_cols = []
    seen = {}
    for i, c in enumerate(df.columns):
        if pd.isna(c) or str(c).strip() == "":
            c_str = f"열_{i+1}"
        else:
            c_str = str(c).strip().replace("\n", " ").replace("\r", "")
        if c_str in seen:
            seen[c_str] += 1
            new_cols.append(f"{c_str}_{seen[c_str]}")
        else:
            seen[c_str] = 0
            new_cols.append(c_str)
    res_df = df.copy()
    res_df.columns = new_cols

    # 숫자 열이 실수(float)로 읽혀 끝에 .0이 붙는 현상 자동 복원
    for col in res_df.columns:
        try:
            s = res_df[col]
            if pd.api.types.is_float_dtype(s):
                non_null = s.dropna()
                if not non_null.empty and (non_null % 1 == 0).all():
                    if (non_null >= -9223372036854775808).all() and (non_null <= 9223372036854775807).all():
                        res_df[col] = s.astype("Int64").astype(str).replace("<NA>", "")
            elif pd.api.types.is_object_dtype(s):
                res_df[col] = s.apply(lambda x: clean_scalar_str(x) if pd.notna(x) else "")
        except Exception:
            pass

    return res_df

def read_uploaded_tabular_file(uploaded_file: Any) -> pd.DataFrame:
    """
    업로드된 엑셀(.xlsx, .xls) 및 CSV, TSV, HTML 테이블 파일을
    다양한 인코딩 및 엔진(openpyxl, xlrd, csv autodetect, read_html)을 시도하여
    오류 없이 안전하고 견고하게 DataFrame으로 변환합니다.

    Args:
        uploaded_file (Any): Streamlit의 UploadedFile 객체, 또는 파일 버퍼/바이트 스트림

    Returns:
        pd.DataFrame: 정상 로드된 Pandas 데이터프레임
    """
    if uploaded_file is None:
        return pd.DataFrame()

    filename: str = getattr(uploaded_file, "name", "").lower()

    # 안전하게 파일 바이트 읽기
    try:
        content_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    except Exception:
        content_bytes = b""

    if not content_bytes or len(content_bytes) == 0:
        raise ValueError("파일 크기가 0바이트(빈 파일)이거나 기기/클라우드에서 다운로드가 완료되지 않았습니다.")

    errors: list[str] = []

    # 1. Excel 엔진 시도 (.xlsx, .xls 또는 확장자 무관 우선 시도)
    # 1-1. xlrd 엔진 우선 시도 (.xls 파일인 경우)
    if filename.endswith(".xls"):
        try:
            return pd.read_excel(io.BytesIO(content_bytes), engine="xlrd")
        except Exception as e:
            errors.append(f"xlrd: {e}")

    # 1-2. openpyxl 엔진 우선 시도 (.xlsx 파일인 경우)
    try:
        return pd.read_excel(io.BytesIO(content_bytes), engine="openpyxl")
    except Exception as e:
        errors.append(f"openpyxl: {e}")

    # 1-3. xlrd 재시도 (.xls 확장자가 아니더라도 구형 바이너리 xls일 가능성 대비)
    try:
        return pd.read_excel(io.BytesIO(content_bytes), engine="xlrd")
    except Exception as e:
        errors.append(f"xlrd_fallback: {e}")

    # 1-4. pandas 기본 엔진 시도
    try:
        return pd.read_excel(io.BytesIO(content_bytes))
    except Exception as e:
        errors.append(f"read_excel_default: {e}")

    # 2. CSV / TSV 텍스트 파일 시도 (확장자는 .xls/.xlsx이지만 실제로는 CSV인 쇼핑몰 발주서 대비)
    encodings = ["utf-8-sig", "cp949", "euc-kr", "utf-8", "latin1"]
    for enc in encodings:
        try:
            df_csv = pd.read_csv(io.BytesIO(content_bytes), encoding=enc, sep=None, engine="python")
            if df_csv is not None and not df_csv.empty and len(df_csv.columns) >= 1:
                return df_csv
        except Exception as e:
            errors.append(f"csv({enc}): {e}")

    # 3. HTML 테이블 시도 (일부 쇼핑몰에서 HTML 테이블 형식을 .xls로 내보내는 경우)
    try:
        dfs = pd.read_html(io.BytesIO(content_bytes))
        if dfs and len(dfs) > 0:
            return dfs[0]
    except Exception as e:
        errors.append(f"read_html: {e}")

    # 모든 시도가 실패했을 때 친절한 예외 메시지
    raise ValueError(
        f"파일을 읽을 수 없습니다. 지원되는 엑셀(.xlsx, .xls) 또는 CSV 파일인지 확인해주세요. (원인: {'; '.join(errors[-2:])})"
    )

def extract_unique_products(df, col_map):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty or not col_map:
        return []
    
    prod_col = col_map.get("product_name")
    opt_col = col_map.get("option_name")
    
    if not prod_col or prod_col == "(선택안함)" or prod_col not in df.columns:
        return []
    
    # 옵션명이 선택안함이거나, df에 없거나, 상품명 컬럼과 동일한 경우 옵션 분리 제외
    has_opt = bool(opt_col and opt_col != "(선택안함)" and opt_col in df.columns and opt_col != prod_col)
    
    try:
        use_cols = [prod_col]
        if has_opt:
            use_cols.append(opt_col)
        
        # 고유 컬럼 목록만 선택하여 Series 복수 반환 원천 차단
        sub_df = df[use_cols].drop_duplicates()
        
        items = []
        for _, row in sub_df.iterrows():
            raw_p = row[prod_col]
            if isinstance(raw_p, (pd.Series, list, tuple, np.ndarray)):
                raw_p = raw_p.iloc[0] if hasattr(raw_p, "iloc") and not raw_p.empty else (raw_p[0] if len(raw_p) > 0 else "")
            p_val = clean_scalar_str(raw_p)
                
            o_val = ""
            if has_opt:
                raw_o = row[opt_col]
                if isinstance(raw_o, (pd.Series, list, tuple, np.ndarray)):
                    raw_o = raw_o.iloc[0] if hasattr(raw_o, "iloc") and not raw_o.empty else (raw_o[0] if len(raw_o) > 0 else "")
                o_val = clean_scalar_str(raw_o)
                    
            if p_val:
                items.append((p_val, o_val))
        return items
    except Exception:
        return []

def process_seller_dataframe(df, seller_name, col_map, product_mapping_dict=None, custom_order_date=None, force_custom_date=False, custom_seller_name=None):
    """
    셀러 데이터프레임 처리:
    - 커스텀 발주 일자 및 판매처(쇼핑몰명) 오버라이드 지원
    - 엑셀 내 판매처 열 지정 시 2행 값 아래로 빈 셀 자동 채우기(ffill)
    - 지정된 본사 단가(unit_price)를 우선 적용하여 주문금액을 정확히 산출
    """
    df = sanitize_dataframe_columns(df)
    norm_data = []

    # 🏪 판매처(쇼핑몰명) 전처리
    custom_seller_clean = str(custom_seller_name).strip() if custom_seller_name else ""
    channel_col = col_map.get("seller_channel", "")
    has_channel_col = bool(channel_col and channel_col != "(선택안함)" and channel_col in df.columns)

    if has_channel_col:
        s_channels = df[channel_col].astype(str).replace(["", "nan", "None", "(없음)"], np.nan).ffill()
    else:
        s_channels = None
    
    clean_fallback_seller = extract_clean_seller_from_filename(seller_name) if seller_name else ""
    default_row_seller = clean_fallback_seller if clean_fallback_seller else str(seller_name or "")
    
    def get_val(row, field_key, default=""):
        col_name = col_map.get(field_key, "")
        if col_name and col_name != "(선택안함)" and col_name in row:
            val = row[col_name]
            if isinstance(val, (pd.Series, list, tuple, np.ndarray)):
                val = val.iloc[0] if hasattr(val, "iloc") and not val.empty else (val[0] if len(val) > 0 else "")
            return clean_scalar_str(val, default)
        return default

    for idx, row in df.iterrows():
        order_time_raw = get_val(row, "order_time")
        order_id = get_val(row, "order_id", f"ORD_{idx+1}")
        if not order_id:
            order_id = f"ORD_{idx+1}"
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

        # 🏪 판매처(쇼핑몰명) 결정: 직접 입력값 > 매칭 열(ffill) > 파일명 정제 셀러명
        if custom_seller_clean:
            row_seller = custom_seller_clean
        elif s_channels is not None and pd.notna(s_channels.iloc[idx]) and str(s_channels.iloc[idx]).strip() and str(s_channels.iloc[idx]).strip() not in ["nan", "None"]:
            row_seller = str(s_channels.iloc[idx]).strip()
        else:
            row_seller = default_row_seller

        # 수량 정수화
        qty_raw = get_val(row, "quantity", "1")
        try:
            qty = int(float(str(qty_raw).replace(",", ""))) if qty_raw else 1
        except Exception:
            qty = 1

        # 상품 및 본사 판매단가 매핑
        std_product = raw_product
        std_option = raw_option
        unit_price = 0

        if product_mapping_dict:
            key_exact = f"{raw_product}||{raw_option}"
            key_prod_only = f"{raw_product}||"
            matched_item = product_mapping_dict.get(key_exact) or product_mapping_dict.get(key_prod_only)
            if matched_item:
                std_product = matched_item.get("brand_product", raw_product)
                std_option = matched_item.get("brand_option", raw_option)
                unit_price = int(matched_item.get("unit_price", 0))

        # 셀러 엑셀에 적힌 금액 (단가가 없을 때 보조로 사용)
        seller_price_raw = get_val(row, "price", "0")
        try:
            seller_price = int(float(str(seller_price_raw).replace(",", ""))) if seller_price_raw else 0
        except Exception:
            seller_price = 0

        # 최종 주문 계산 금액: 본사 단가가 설정되어 있으면 (단가 * 수량), 없으면 셀러 엑셀 금액 사용
        if unit_price > 0:
            calc_price = unit_price * qty
        else:
            calc_price = seller_price

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
                    det_date = extract_date_from_filename(seller_name)
                    order_date = det_date.strftime("%Y-%m-%d")
                    try:
                        parsed_time = pd.Timestamp(order_date)
                    except Exception:
                        parsed_time = pd.Timestamp.now()
            else:
                order_date = parsed_time.strftime("%Y-%m-%d")

            order_time_str = parsed_time.strftime("%Y-%m-%d %H:%M:%S") if order_time_raw else f"{order_date} 00:00:00"

        full_addr = f"{address} {address_detail}".strip()
        is_jeju = is_jeju_island(full_addr, zipcode)

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
            "unit_price": unit_price,
            "quantity": qty,
            "price": calc_price,
            "shipping_memo": shipping_memo,
            "is_jeju": is_jeju,
            "event_tags": [],
            "gifts_to_add": []
        })

    return pd.DataFrame(norm_data)

def apply_event_rules(df, event_config):
    if df.empty:
        return df

    n_rows = len(df)
    row_event_tags = [[] for _ in range(n_rows)]
    row_applied_events = [[] for _ in range(n_rows)]
    row_gifts = [[] for _ in range(n_rows)]
    df = df.copy()

    # 식별자 추출 헬퍼 (주문번호 / 수취인명 / 연락처 / 주소 체크리스트 완벽 지원 ⭐)
    def extract_entity_id(row, rule_spec):
        keys = []
        if isinstance(rule_spec, (list, set, tuple)):
            keys = list(rule_spec)
        elif isinstance(rule_spec, str):
            tokens = []
            if "주문번호" in rule_spec: tokens.append("주문번호")
            if "수취인명" in rule_spec or "수령인" in rule_spec or "구매자" in rule_spec: tokens.append("수취인명")
            if "연락처" in rule_spec or "전화" in rule_spec: tokens.append("연락처")
            if "주소" in rule_spec: tokens.append("주소")
            if tokens:
                keys = tokens
            elif "동일 구매자" in rule_spec or "1인 1회" in rule_spec:
                keys = ["수취인명", "연락처"]
            else:
                keys = ["주문번호"]

        if not keys:
            keys = ["수취인명", "연락처"]

        parts = []
        for k in keys:
            if k == "주문번호":
                parts.append(str(row.get("order_id", "")).strip())
            elif k in ["수취인명", "구매자", "수령인"]:
                clean_name = str(row.get("receiver_name") or row.get("buyer_name") or "").strip()
                parts.append(clean_name)
            elif k in ["연락처", "전화"]:
                clean_phone = re.sub(r"[^0-9]", "", str(row.get("receiver_phone") or row.get("buyer_phone") or ""))
                parts.append(clean_phone)
            elif k in ["주소", "배송지주소"]:
                clean_addr = str(row.get("full_address") or row.get("address") or "").strip()
                clean_addr = re.sub(r"\s+", " ", clean_addr)
                parts.append(clean_addr)

        return "___".join(parts) if parts else str(row.get("order_id", ""))

    # 1. 선착순 이벤트 (독립 규칙 적용)
    fcfs_cfg = event_config.get("first_come_first_served", {})
    if fcfs_cfg.get("enabled", False):
        limit = fcfs_cfg.get("limit", 3)
        gift_prod = fcfs_cfg.get("gift_product_name", "[증정] 선착순 사은품")
        gift_opt = fcfs_cfg.get("gift_option_name", "증정용")
        gift_qty = fcfs_cfg.get("gift_quantity", 1)
        target_rule = fcfs_cfg.get("target_rule", "동일 구매자(수취인 연락처) 기준 (1인 1회 한정 ⭐)")
        display_rule = fcfs_cfg.get("tag_display_rule", "텍스트 표기 없음 (화면 및 엑셀 배경색 하이라이트 색인만 처리 ⭐)")
        gift_pos_rule = fcfs_cfg.get("gift_position", "해당 주문건의 맨 마지막 행에 추가 (추천 ⭐)")
        action_type = fcfs_cfg.get("action_type", "add_row")
        fcfs_no_color = fcfs_cfg.get("no_color", False)
        fcfs_hl_scope = fcfs_cfg.get("highlight_scope", "both")
        fcfs_write_col = fcfs_cfg.get("write_col_enabled", False)
        fcfs_target_col = fcfs_cfg.get("write_target_col", "N")

        sorted_df = df.sort_values(by="order_time_obj")
        
        winners_set = set()
        seen_entities = []
        for _, s_row in sorted_df.iterrows():
            ent_id = extract_entity_id(s_row, target_rule)
            if ent_id and ent_id not in seen_entities:
                seen_entities.append(ent_id)
                if len(seen_entities) <= limit:
                    winners_set.add(ent_id)
                else:
                    break

        handled_entities = set()
        for i, (idx, row) in enumerate(df.iterrows()):
            ent_id = extract_entity_id(row, target_rule)
            if ent_id in winners_set:
                fcfs_clr = fcfs_cfg.get("color", event_config.get("color_event", "#FFF2CC"))
                tag_str = f"[선착순{limit}명]"
                row_event_tags[i].append(tag_str)
                row_applied_events[i].append({
                    "event_type": "fcfs",
                    "tag": tag_str,
                    "tag_display_rule": display_rule,
                    "write_col_enabled": fcfs_write_col,
                    "write_target_col": fcfs_target_col,
                    "reason": f"선착순 {limit}명 사은품",
                    "color": fcfs_clr,
                    "no_color": fcfs_no_color,
                    "highlight_scope": fcfs_hl_scope
                })
                if action_type == "add_row" and "행 추가 안 함" not in gift_pos_rule:
                    if ent_id not in handled_entities:
                        row_gifts[i].append({
                            "product_name": gift_prod,
                            "option_name": gift_opt,
                            "quantity": gift_qty,
                            "reason": f"선착순 {limit}명 사은품",
                            "order_id": row["order_id"],
                            "gift_position": gift_pos_rule,
                            "write_col_enabled": fcfs_write_col,
                            "write_target_col": fcfs_target_col,
                            "color": fcfs_clr,
                            "no_color": fcfs_no_color,
                            "highlight_scope": fcfs_hl_scope
                        })
                        handled_entities.add(ent_id)

    # 2. 금액 조건 이벤트 (독립 규칙 적용 + 선착순 한정 지원 ⭐)
    amt_cfg = event_config.get("min_amount", {})
    if amt_cfg.get("enabled", False):
        min_amt = amt_cfg.get("min_amount", 50000)
        amt_clr = amt_cfg.get("color", event_config.get("color_event", "#FFF2CC"))
        gift_prod = amt_cfg.get("gift_product_name", "[증정] 금액달성 사은품")
        gift_opt = amt_cfg.get("gift_option_name", "증정용")
        base_gift_qty = amt_cfg.get("gift_quantity", 1)
        target_rule = amt_cfg.get("target_rule", ["주문번호"])
        display_rule = amt_cfg.get("tag_display_rule", "텍스트 표기 없음 (화면 및 엑셀 배경색 하이라이트 색인만 처리 ⭐)")
        gift_pos_rule = amt_cfg.get("gift_position", "해당 주문(합포장) 가장 아래에 사은품 행 추가 (추천 ⭐)")
        amt_fcfs_on = amt_cfg.get("fcfs_enabled", False)
        amt_fcfs_limit = amt_cfg.get("fcfs_limit", 50)
        amt_no_color = amt_cfg.get("no_color", False)
        amt_hl_scope = amt_cfg.get("highlight_scope", "both")
        amt_write_col = amt_cfg.get("write_col_enabled", False)
        amt_target_col = amt_cfg.get("write_target_col", "N")
        
        df["_amt_group_key"] = df.apply(lambda r: extract_entity_id(r, target_rule), axis=1)
        group_totals = df.groupby("_amt_group_key")["price"].sum().to_dict()

        amt_winners = set()
        if amt_fcfs_on:
            qual_groups = {}
            for _, r_tmp in df.iterrows():
                gk = r_tmp["_amt_group_key"]
                if group_totals.get(gk, 0) >= min_amt:
                    ot = r_tmp["order_time_obj"]
                    if gk not in qual_groups or ot < qual_groups[gk]:
                        qual_groups[gk] = ot
            sorted_qual = sorted(qual_groups.items(), key=lambda x: x[1])
            amt_winners = set(k for k, _ in sorted_qual[:amt_fcfs_limit])

        handled_amt_groups = set()
        for i, (idx, row) in enumerate(df.iterrows()):
            g_key = row["_amt_group_key"]
            tot_price = group_totals.get(g_key, 0)
            if tot_price >= min_amt and (not amt_fcfs_on or g_key in amt_winners):
                tag_str = f"[{min_amt:,}원 선착순{amt_fcfs_limit}명]" if amt_fcfs_on else f"[{min_amt:,}원이상]"
                if tag_str not in row_event_tags[i]:
                    row_event_tags[i].append(tag_str)
                row_applied_events[i].append({
                    "event_type": "amt",
                    "tag": tag_str,
                    "tag_display_rule": display_rule,
                    "write_col_enabled": amt_write_col,
                    "write_target_col": amt_target_col,
                    "reason": "금액달성 사은품",
                    "color": amt_clr,
                    "no_color": amt_no_color,
                    "highlight_scope": amt_hl_scope
                })
                
                if "행 추가 안 함" not in gift_pos_rule and "별도 행 추가 안 함" not in gift_pos_rule and g_key not in handled_amt_groups:
                    calc_qty = base_gift_qty
                    if any("배수" in str(x) for x in (target_rule if isinstance(target_rule, list) else [target_rule])):
                        multiplier = max(1, int(tot_price // min_amt))
                        calc_qty = base_gift_qty * multiplier

                    row_gifts[i].append({
                        "product_name": gift_prod,
                        "option_name": gift_opt,
                        "quantity": calc_qty,
                        "reason": "금액달성 사은품",
                        "order_id": row["order_id"],
                        "gift_position": gift_pos_rule,
                        "tag_display_rule": display_rule,
                        "write_col_enabled": amt_write_col,
                        "write_target_col": amt_target_col,
                        "color": amt_clr,
                        "no_color": amt_no_color,
                        "highlight_scope": amt_hl_scope
                    })
                    handled_amt_groups.add(g_key)

    # 3. 특정 품목 구매 사은품 (독립 규칙 적용)
    spec_cfg = event_config.get("special_product", {})
    if spec_cfg.get("enabled", False):
        target_kw = spec_cfg.get("keyword", "")
        gift_prod = spec_cfg.get("gift_product_name", "[사은품] 특정품목 증정")
        gift_opt = spec_cfg.get("gift_option_name", "단품")
        base_gift_qty = spec_cfg.get("gift_quantity", 1)
        target_rule = spec_cfg.get("target_rule", ["주문번호"])
        display_rule = spec_cfg.get("tag_display_rule", "텍스트 표기 없음 (화면 및 엑셀 배경색 하이라이트 색인만 처리 ⭐)")
        gift_pos_rule = spec_cfg.get("gift_position", "해당 주문(합포장) 가장 아래에 사은품 행 추가 (추천 ⭐)")
        spec_no_color = spec_cfg.get("no_color", False)
        spec_hl_scope = spec_cfg.get("highlight_scope", "both")
        spec_write_col = spec_cfg.get("write_col_enabled", False)
        spec_target_col = spec_cfg.get("write_target_col", "N")

        spec_clr = spec_cfg.get("color", event_config.get("color_gift", "#E2EFDA"))
        if target_kw:
            handled_spec_groups = set()
            for i, (idx, row) in enumerate(df.iterrows()):
                target_str = f"{row.get('product_name','')} {row.get('option_name','')} {row.get('raw_product','')}"
                if target_kw in target_str:
                    tag_str = f"[{target_kw}구매]"
                    if tag_str not in row_event_tags[i]:
                        row_event_tags[i].append(tag_str)
                    row_applied_events[i].append({
                        "event_type": "spec",
                        "tag": tag_str,
                        "tag_display_rule": display_rule,
                        "write_col_enabled": spec_write_col,
                        "write_target_col": spec_target_col,
                        "reason": f"{target_kw} 구매 사은품",
                        "color": spec_clr,
                        "no_color": spec_no_color,
                        "highlight_scope": spec_hl_scope
                    })

                    if "행 추가 안 함" not in gift_pos_rule and "별도 행 추가 안 함" not in gift_pos_rule:
                        ent_id = extract_entity_id(row, target_rule)
                        if any("1:1" in str(x) for x in (target_rule if isinstance(target_rule, list) else [target_rule])):
                            row_gifts[i].append({
                                "product_name": gift_prod,
                                "option_name": gift_opt,
                                "quantity": base_gift_qty * row.get("quantity", 1),
                                "reason": f"{target_kw} 구매 사은품",
                                "order_id": row["order_id"],
                                "gift_position": gift_pos_rule,
                                "tag_display_rule": display_rule,
                                "write_col_enabled": spec_write_col,
                                "write_target_col": spec_target_col,
                                "color": spec_clr,
                                "no_color": spec_no_color,
                                "highlight_scope": spec_hl_scope
                            })
                        else:
                            if ent_id not in handled_spec_groups:
                                row_gifts[i].append({
                                    "product_name": gift_prod,
                                    "option_name": gift_opt,
                                    "quantity": base_gift_qty,
                                    "reason": f"{target_kw} 구매 사은품",
                                    "order_id": row["order_id"],
                                    "gift_position": gift_pos_rule,
                                    "tag_display_rule": display_rule,
                                    "write_col_enabled": spec_write_col,
                                    "write_target_col": spec_target_col,
                                    "color": spec_clr,
                                    "no_color": spec_no_color,
                                    "highlight_scope": spec_hl_scope
                                })
                                handled_spec_groups.add(ent_id)

    # 4. 자유 추가 커스텀 사은품 이벤트 (custom_events ⭐)
    custom_events = event_config.get("custom_events", [])
    for c_idx, c_ev in enumerate(custom_events):
        if not c_ev.get("enabled", False):
            continue
        c_title = c_ev.get("title", f"추가행사{c_idx+1}")
        c_cond = c_ev.get("condition_type", "전체 주문")
        c_gift_prod = c_ev.get("gift_product_name", "[증정] 추가 사은품")
        c_gift_opt = c_ev.get("gift_option_name", "단품")
        c_gift_qty = c_ev.get("gift_quantity", 1)
        c_target_rule = c_ev.get("target_rule", ["주문번호"])
        c_pos_rule = c_ev.get("gift_position", "해당 주문(합포장) 가장 아래에 사은품 행 추가 (추천 ⭐)")
        c_disp_rule = c_ev.get("tag_display_rule", "텍스트 표기 없음 (화면 및 엑셀 배경색 하이라이트 색인만 처리 ⭐)")
        c_clr = c_ev.get("color", event_config.get("color_event", "#FFF2CC"))
        c_no_color = c_ev.get("no_color", False)
        c_hl_scope = c_ev.get("highlight_scope", "both")
        c_write_col = c_ev.get("write_col_enabled", False)
        c_target_col = c_ev.get("write_target_col", "N")

        c_qualifying_indices = []
        if "선착순" in c_cond:
            c_limit = c_ev.get("limit", 10)
            sorted_cdf = df.sort_values(by="order_time_obj")
            c_winners = set()
            c_seen = []
            for _, s_row in sorted_cdf.iterrows():
                ent = extract_entity_id(s_row, c_target_rule)
                if ent and ent not in c_seen:
                    c_seen.append(ent)
                    if len(c_seen) <= c_limit:
                        c_winners.add(ent)
                    else:
                        break
            for i, (idx, row) in enumerate(df.iterrows()):
                ent = extract_entity_id(row, c_target_rule)
                if ent in c_winners:
                    c_qualifying_indices.append(i)

        elif "금액" in c_cond:
            c_min_amt = c_ev.get("min_amount", 30000)
            col_tmp_key = f"_c_amt_{c_idx}"
            df[col_tmp_key] = df.apply(lambda r: extract_entity_id(r, c_target_rule), axis=1)
            c_totals = df.groupby(col_tmp_key)["price"].sum().to_dict()
            for i, (idx, row) in enumerate(df.iterrows()):
                ent = row[col_tmp_key]
                if c_totals.get(ent, 0) >= c_min_amt:
                    c_qualifying_indices.append(i)
            if col_tmp_key in df.columns:
                df.drop(columns=[col_tmp_key], inplace=True)

        elif "특정 품목" in c_cond or "키워드" in c_cond:
            c_kw = c_ev.get("keyword", "")
            if c_kw:
                for i, (idx, row) in enumerate(df.iterrows()):
                    target_str = f"{row.get('product_name','')} {row.get('option_name','')} {row.get('raw_product','')}"
                    if c_kw in target_str:
                        c_qualifying_indices.append(i)

        else: # "전체 주문 증정"
            c_qualifying_indices = list(range(len(df)))

        c_handled_entities = set()
        for i in c_qualifying_indices:
            row = df.iloc[i]
            tag_str = f"[{c_title}]"
            if tag_str not in row_event_tags[i]:
                row_event_tags[i].append(tag_str)
            row_applied_events[i].append({
                "event_type": "custom",
                "tag": tag_str,
                "tag_display_rule": c_disp_rule,
                "write_col_enabled": c_write_col,
                "write_target_col": c_target_col,
                "reason": c_title,
                "color": c_clr,
                "no_color": c_no_color,
                "highlight_scope": c_hl_scope
            })
            if "행 추가 안 함" not in c_pos_rule and "별도 행 추가 안 함" not in c_pos_rule:
                ent = extract_entity_id(row, c_target_rule)
                if ent not in c_handled_entities:
                    row_gifts[i].append({
                        "product_name": c_gift_prod,
                        "option_name": c_gift_opt,
                        "quantity": c_gift_qty,
                        "reason": c_title,
                        "order_id": row["order_id"],
                        "gift_position": c_pos_rule,
                        "tag_display_rule": c_disp_rule,
                        "write_col_enabled": c_write_col,
                        "write_target_col": c_target_col,
                        "color": c_clr,
                        "no_color": c_no_color,
                        "highlight_scope": c_hl_scope
                    })
                    c_handled_entities.add(ent)

    # 임시 컬럼 정리
    if "_amt_group_key" in df.columns:
        df.drop(columns=["_amt_group_key"], inplace=True)

    df["event_tags"] = row_event_tags
    df["applied_events"] = row_applied_events
    df["gifts_to_add"] = row_gifts
    return df

def col_letter_to_index(col_str):
    if not col_str:
        return None
    s = str(col_str).strip().upper()
    if s.isalpha():
        num = 0
        for c in s:
            num = num * 26 + (ord(c) - ord('A') + 1)
        return num - 1
    return None

def index_to_col_letter(idx):
    result = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

def resolve_or_add_target_column(brand_cols, col_spec):
    col_spec_clean = str(col_spec).strip()
    if not col_spec_clean:
        col_spec_clean = "N"

    for col in brand_cols:
        if col.strip().lower() == col_spec_clean.lower():
            return col

    idx = col_letter_to_index(col_spec_clean)
    if idx is not None:
        letter = col_spec_clean.upper()
        if idx < len(brand_cols):
            existing_col = brand_cols[idx]
            norm = existing_col.replace(" ", "").lower()
            if any(k in norm for k in ["비고", "이벤트", "색인", "특이사항", "메모", "note", "remark"]):
                return existing_col
            col_name = f"비고({letter})"
            if col_name in brand_cols:
                return col_name
            brand_cols.insert(idx, col_name)
            return col_name
        else:
            while len(brand_cols) < idx:
                curr_letter = index_to_col_letter(len(brand_cols))
                brand_cols.append(f"공란({curr_letter})")
            col_name = f"비고({letter})"
            brand_cols.append(col_name)
            return col_name
    else:
        col_name = col_spec_clean
        if col_name not in brand_cols:
            brand_cols.append(col_name)
        return col_name

def build_brand_order_sheet(df, brand_template, event_config=None):
    if event_config is None:
        event_config = {}

    brand_cols = list(brand_template.get("columns", [
        "순번", "주문일시", "주문번호", "수령인", "연락처", "우편번호", "주소", "상세주소",
        "브랜드상품명", "옵션명", "수량", "배송메모", "제주도서여부", "이벤트색인"
    ]))

    # 사용자 지정 비고열(영문 열 이니셜 or 컬럼명) 사전 등록 및 매핑
    target_col_map = {}
    active_ev_configs = []
    if event_config.get("first_come_first_served", {}).get("enabled"):
        active_ev_configs.append(event_config["first_come_first_served"])
    if event_config.get("min_amount", {}).get("enabled"):
        active_ev_configs.append(event_config["min_amount"])
    if event_config.get("special_product", {}).get("enabled"):
        active_ev_configs.append(event_config["special_product"])
    for cev in event_config.get("custom_events", []):
        if cev.get("enabled"):
            active_ev_configs.append(cev)

    for ev_c in active_ev_configs:
        if ev_c.get("write_col_enabled"):
            cspec = str(ev_c.get("write_target_col", "N")).strip()
            if cspec and cspec not in target_col_map:
                resolved_col = resolve_or_add_target_column(brand_cols, cspec)
                target_col_map[cspec] = resolved_col

    rows = []
    seq = 1
    end_of_sheet_gifts = []

    # 각 주문번호별 마지막 행 인덱스 및 사은품 모음 사전 구축
    last_row_of_order = {}
    order_gifts_map = {}
    for idx, r in df.iterrows():
        oid = r["order_id"]
        last_row_of_order[oid] = idx
        if r.get("gifts_to_add"):
            if oid not in order_gifts_map:
                order_gifts_map[oid] = []
            order_gifts_map[oid].extend(r["gifts_to_add"])

    for idx, r in df.iterrows():
        # 각 프로모션별 독립 표기 규칙 평가
        biko_event_tags = []
        memo_event_tags = []
        prod_event_tags = []
        has_highlight = False

        applied_evs = r.get("applied_events", [])
        for ev in applied_evs:
            disp_rule = ev.get("tag_display_rule", "")
            tag_val = ev.get("tag", "")

            if "완전 미표기" in disp_rule:
                continue

            has_highlight = True

            if "텍스트 표기 없음" in disp_rule or "색인만" in disp_rule:
                pass # 텍스트 표기 생략 (배경색 하이라이트만 작동)
            elif "배송메세지" in disp_rule and "비고 +" not in disp_rule:
                memo_event_tags.append(tag_val)
            elif "비고 + 배송메세지" in disp_rule:
                biko_event_tags.append(tag_val)
                memo_event_tags.append(tag_val)
            elif "상품명 뒤에" in disp_rule:
                prod_event_tags.append(tag_val)
            else: # "비고 / 이벤트색인 열에만 표기"
                biko_event_tags.append(tag_val)

        # 텍스트 합성 (기존 원본 데이터 보존)
        biko_event = " / ".join(biko_event_tags)
        memo_raw = r.get("shipping_memo", "")
        if memo_event_tags:
            m_tag_str = " / ".join(memo_event_tags)
            memo_event = f"{memo_raw} ({m_tag_str})".strip(" ()") if memo_raw else m_tag_str
        else:
            memo_event = memo_raw

        prod_raw = r.get("product_name", "")
        if prod_event_tags:
            prod_event = f"{prod_raw} ({' / '.join(prod_event_tags)})".strip()
        else:
            prod_event = prod_raw

        field_pool = {
            "순번": seq,
            "쇼핑몰명": r["seller_name"],
            "셀러명": r["seller_name"],
            "판매처": r["seller_name"],
            "판매처명": r["seller_name"],
            "몰명": r["seller_name"],
            "채널명": r["seller_name"],
            "날짜": r["order_date"],
            "주문일자": r["order_date"],
            "발주일자": r["order_date"],
            "발주일": r["order_date"],
            "주문일시": r["order_time"],
            "결제일시": r["order_time"],
            "쇼핑몰 주문번호": clean_scalar_str(r["order_id"]),
            "쇼핑몰주문번호": clean_scalar_str(r["order_id"]),
            "사방넷 주문번호": clean_scalar_str(r["order_id"]),
            "사방넷주문번호": clean_scalar_str(r["order_id"]),
            "주문번호": clean_scalar_str(r["order_id"]),
            "주문번호.1": clean_scalar_str(r["order_id"]),
            "주문번호.2": clean_scalar_str(r["order_id"]),
            "발주번호": clean_scalar_str(r["order_id"]),
            "구매자": r["buyer_name"],
            "구매자명": r["buyer_name"],
            "주문자": r["buyer_name"],
            "주문자명": r["buyer_name"],
            "구매자연락처": clean_scalar_str(r["buyer_phone"]),
            "주문자연락처": clean_scalar_str(r["buyer_phone"]),
            "구매자전화번호": clean_scalar_str(r["buyer_phone"]),
            "주문자전화번호": clean_scalar_str(r["buyer_phone"]),
            "수취인명": r["receiver_name"],
            "수령인": r["receiver_name"],
            "수령인명": r["receiver_name"],
            "수령자": r["receiver_name"],
            "수령자명": r["receiver_name"],
            "수취인": r["receiver_name"],
            "받는분": r["receiver_name"],
            "받으시는분": r["receiver_name"],
            "받는사람": r["receiver_name"],
            "고객명": r["receiver_name"],
            "수취인전화번호1": clean_scalar_str(r["receiver_phone"]),
            "수취인전화번호": clean_scalar_str(r["receiver_phone"]),
            "수취인핸드폰": clean_scalar_str(r["receiver_phone"]),
            "수취인휴대폰": clean_scalar_str(r["receiver_phone"]),
            "수취인연락처": clean_scalar_str(r["receiver_phone"]),
            "수령자전화번호": clean_scalar_str(r["receiver_phone"]),
            "수령자전화": clean_scalar_str(r["receiver_phone"]),
            "수령자핸드폰": clean_scalar_str(r["receiver_phone"]),
            "수령자휴대폰": clean_scalar_str(r["receiver_phone"]),
            "수령자연락처": clean_scalar_str(r["receiver_phone"]),
            "수령인전화번호": clean_scalar_str(r["receiver_phone"]),
            "수령인전화": clean_scalar_str(r["receiver_phone"]),
            "수령인핸드폰": clean_scalar_str(r["receiver_phone"]),
            "수령인휴대폰": clean_scalar_str(r["receiver_phone"]),
            "수령인연락처": clean_scalar_str(r["receiver_phone"]),
            "연락처": clean_scalar_str(r["receiver_phone"]),
            "전화번호": clean_scalar_str(r.get("receiver_phone", "")),
            "휴대폰": clean_scalar_str(r.get("receiver_phone", "")),
            "핸드폰": clean_scalar_str(r.get("receiver_phone", "")),
            "우편번호": clean_scalar_str(r.get("zipcode", "")),
            "수령자주소": r.get("full_address") or r.get("address", ""),
            "수령인주소": r.get("full_address") or r.get("address", ""),
            "수취인주소": r.get("full_address") or r.get("address", ""),
            "배송지주소": r.get("full_address") or r.get("address", ""),
            "배송주소": r.get("full_address") or r.get("address", ""),
            "전체주소": r.get("full_address") or r.get("address", ""),
            "기본주소": r.get("address", ""),
            "주소": r.get("address", ""),
            "주소1": r.get("address", ""),
            "상세주소": r.get("address_detail", ""),
            "주소2": r.get("address_detail", ""),
            "배송지상세": r.get("address_detail", ""),
            "상품약어": prod_event,
            "브랜드상품명": prod_event,
            "상품명": prod_event,
            "사방넷상품코드": "",
            "옵션코드 기입": r.get("option_name", ""),
            "옵션명": r.get("option_name", ""),
            "옵션": r.get("option_name", ""),
            "수량": r.get("quantity", 1),
            "결제금액": r.get("price", 0),
            "판매가": r.get("price", 0),
            "배송메세지": memo_event,
            "배송메모": memo_event,
            "물류처명": "기본물류",
            "송장번호": "",
            "운송장번호": "",
            "제주도서여부": "제주/도서산간" if r.get("is_jeju") else "일반",
            "이벤트색인": biko_event,
            "비고": biko_event
        }

        def resolve_brand_column_value(col, r, pool, cur_seq=1):
            if col in pool and pool[col] is not None and str(pool[col]).strip() != "":
                return pool[col]

            clean_col = re.sub(r'\.\d+$', '', str(col)).strip()
            if clean_col in pool and pool[clean_col] is not None and str(pool[clean_col]).strip() != "":
                return pool[clean_col]

            norm = clean_col.replace(" ", "").lower()

            # [1순위] 전화번호 / 핸드폰 / 휴대폰 / 연락처 (반드시 사람 이름보다 먼저!)
            if any(k in norm for k in ["전화", "핸드폰", "휴대폰", "연락처", "phone", "tel", "hp"]):
                if any(k in norm for k in ["주문자", "구매자", "buyer"]):
                    return r.get("buyer_phone", "")
                return r.get("receiver_phone", "")

            # [2순위] 주소 / 배송지 (반드시 사람 이름보다 먼저!)
            if any(k in norm for k in ["주소", "배송지", "배송주소", "address"]):
                if any(k in norm for k in ["상세", "주소2", "detail"]):
                    return r.get("address_detail", "")
                return r.get("full_address") or r.get("address", "")

            # [3순위] 우편번호
            if any(k in norm for k in ["우편", "zipcode", "postcode"]):
                return r.get("zipcode", "")

            # [4순위] 배송메모 / 요청사항
            if any(k in norm for k in ["배송메세지", "배송메모", "배송요청", "요청사항", "메모", "memo"]):
                return pool.get("배송메세지") or r.get("shipping_memo", "")

            # [5순위] 상품명 / 품목명
            if any(k in norm for k in ["상품", "품목", "약어", "브랜드상품명", "product"]):
                return pool.get("브랜드상품명") or pool.get("상품명") or r.get("product_name", "")

            # [6순위] 옵션
            if any(k in norm for k in ["옵션", "규격", "option"]):
                return pool.get("옵션명") or pool.get("옵션코드 기입") or r.get("option_name", "")

            # [7순위] 수량
            if any(k in norm for k in ["수량", "qty", "개수", "구매수량"]):
                return pool.get("수량", r.get("quantity", 1))

            # [8순위] 결제금액 / 판매가
            if any(k in norm for k in ["결제금액", "판매가", "단가", "주문금액", "금액", "price"]):
                return pool.get("결제금액", r.get("price", 0))

            # [9순위] 날짜 / 일시
            if any(k in norm for k in ["날짜", "일자", "주문일", "결제일", "발주일", "출고일", "주문일시", "date", "time"]):
                return r.get("order_date") or r.get("order_time", "")

            # [10순위] 주문번호
            if any(k in norm for k in ["주문번호", "쇼핑몰주문번호", "사방넷주문번호", "발주번호", "주문코드", "order_id"]):
                return r.get("order_id", "")

            # [11순위] 판매처 / 셀러명
            if any(k in norm for k in ["판매처", "쇼핑몰", "쇼핑몰명", "셀러명", "판매몰", "거래처", "업체명", "사이트", "마켓", "스토어", "공구처"]):
                return r.get("seller_name", "")

            # [12순위] 사람 이름 (전화번호, 주소가 다 걸러진 후 순수한 이름만!)
            if any(k in norm for k in ["주문자", "주문자명", "구매자", "구매자명", "입금자", "보내는분"]):
                return r.get("buyer_name", "")

            if any(k in norm for k in ["수령자", "수령인", "수취인", "수취인명", "수령인명", "수령자명", "받는분", "받으시는분", "받는사람", "고객명", "수하인"]):
                return r.get("receiver_name", "")

            # [13순위] 순번 / 송장번호 / 제주 / 비고
            if norm in ["순번", "번호", "no", "num"]:
                return cur_seq

            if any(k in norm for k in ["송장번호", "운송장번호", "송장"]):
                return ""

            if any(k in norm for k in ["제주", "도서산간"]):
                return "제주/도서산간" if r.get("is_jeju") else "일반"

            if any(k in norm for k in ["이벤트색인", "비고", "특이사항"]):
                return pool.get("비고", "")

            return pool.get(col, "")

        # 본품 행 추가
        row_dict = {}
        for col in brand_cols:
            row_dict[col] = resolve_brand_column_value(col, r, field_pool, seq)

        # 사용자 지정 특정 열(비고열)에 이벤트 태그 기재
        col_tags_map = {}
        for ev in applied_evs:
            if ev.get("write_col_enabled"):
                cspec = str(ev.get("write_target_col", "N")).strip()
                res_col = target_col_map.get(cspec) or resolve_or_add_target_column(brand_cols, cspec)
                if res_col not in col_tags_map:
                    col_tags_map[res_col] = []
                t_str = ev.get("tag", "")
                if t_str and t_str not in col_tags_map[res_col]:
                    col_tags_map[res_col].append(t_str)

        for res_col, tags in col_tags_map.items():
            if res_col in brand_cols:
                existing_val = str(row_dict.get(res_col, "")).strip()
                tag_combined = " / ".join(tags)
                if existing_val:
                    row_dict[res_col] = f"{existing_val} {tag_combined}".strip()
                else:
                    row_dict[res_col] = tag_combined

        # 본품 행 색인 여부 및 색상 판정
        apply_item_hl = False
        item_color = ""
        for ev in applied_evs:
            ev_no_color = ev.get("no_color", False)
            ev_scope = ev.get("highlight_scope", "both")
            # 색인 처리 안 함이 아니고, "사은품 행에만 색인"이 아닌 경우(즉 "both")에만 본품에 색인!
            if not ev_no_color and ev_scope != "gift_only":
                apply_item_hl = True
                if ev.get("color"):
                    item_color = ev.get("color")
                    break

        if apply_item_hl:
            row_dict["_row_highlight"] = "event"
            row_dict["_row_color"] = item_color or event_config.get("color_event", "#FFF2CC")
        else:
            row_dict["_row_highlight"] = "none"
            row_dict["_row_color"] = ""

        rows.append(row_dict)
        seq += 1

        def make_gift_row(gift_item, base_pool):
            nonlocal seq
            g_pool = base_pool.copy()
            g_pool["순번"] = seq
            g_pool["상품약어"] = gift_item["product_name"]
            g_pool["브랜드상품명"] = gift_item["product_name"]
            g_pool["상품명"] = gift_item["product_name"]
            g_pool["옵션코드 기입"] = gift_item.get("option_name", "단품")
            g_pool["옵션명"] = gift_item.get("option_name", "단품")
            g_pool["옵션"] = gift_item.get("option_name", "단품")
            g_pool["수량"] = gift_item.get("quantity", 1)
            g_pool["결제금액"] = 0
            g_pool["판매가"] = 0
            g_pool["이벤트색인"] = f"[사은품] {gift_item['reason']}"
            g_pool["비고"] = f"[사은품] {gift_item['reason']}"

            # 배송메세지는 원본 고객 배송메세지 유지, 사용자가 명시적으로 '배송메세지' 표기를 선택했을 때만 덧붙임!
            raw_memo = r.get("shipping_memo", "")
            g_disp_rule = gift_item.get("tag_display_rule", "")
            if "배송메세지" in g_disp_rule:
                g_pool["배송메세지"] = f"{raw_memo} (본품합포장: {gift_item['reason']})".strip() if raw_memo else f"본품합포장: {gift_item['reason']}"
                g_pool["배송메모"] = g_pool["배송메세지"]
            else:
                g_pool["배송메세지"] = raw_memo
                g_pool["배송메모"] = raw_memo
            seq += 1

            gift_dict = {}
            for col in brand_cols:
                gift_dict[col] = resolve_brand_column_value(col, r, g_pool, seq - 1)
            
            # 사은품 행에도 특정 열 기재 옵션이 켜져있으면 기재
            if gift_item.get("write_col_enabled"):
                cspec = str(gift_item.get("write_target_col", "N")).strip()
                res_col = target_col_map.get(cspec) or resolve_or_add_target_column(brand_cols, cspec)
                if res_col in brand_cols:
                    gift_tag = f"[사은품] {gift_item.get('reason', '')}"
                    existing_val = str(gift_dict.get(res_col, "")).strip()
                    if existing_val:
                        gift_dict[res_col] = f"{existing_val} {gift_tag}".strip()
                    else:
                        gift_dict[res_col] = gift_tag

            g_no_color = gift_item.get("no_color", False)
            if g_no_color:
                gift_dict["_row_highlight"] = "none"
                gift_dict["_row_color"] = ""
            else:
                gift_dict["_row_highlight"] = "gift"
                gift_dict["_row_color"] = gift_item.get("color") or event_config.get("color_gift", "#E2EFDA")
            return gift_dict

        # 사은품 분기 처리
        # (1) "본품 바로 아래" 모드인 사은품은 현재 본품 행 바로 뒤에 출력
        if r.get("gifts_to_add"):
            for g in r["gifts_to_add"]:
                pos_mode = g.get("gift_position", "")
                if "본품 바로 아래" in pos_mode:
                    rows.append(make_gift_row(g, field_pool))
                elif "최하단" in pos_mode or "맨 끝" in pos_mode or "모아서" in pos_mode:
                    end_of_sheet_gifts.append((g, field_pool))

        # (2) "해당 주문(합포장) 가장 아래" 모드:
        # 이 주문의 맨 마지막 품목 행에 도달했을 때, 이 주문에 배정된 모든 합포장 사은품을 누락 없이 일괄 추가!
        # (단일 품목 구매자뿐만 아니라 여러 품목 장바구니 구매자도 100% 마지막 행에 정상 추가됨)
        if last_row_of_order.get(r["order_id"]) == idx:
            order_gifts = order_gifts_map.get(r["order_id"], [])
            for g in order_gifts:
                pos_mode = g.get("gift_position", "해당 주문(합포장) 가장 아래에 사은품 행 추가 (추천 ⭐)")
                if "행 추가 안 함" in pos_mode or "행추가 안 함" in pos_mode or "별도 행 추가 안 함" in pos_mode:
                    continue
                if "본품 바로 아래" in pos_mode or "최하단" in pos_mode or "맨 끝" in pos_mode or "모아서" in pos_mode:
                    continue
                rows.append(make_gift_row(g, field_pool))

    # 발주서 맨 끝에 모아서 추가된 사은품 행 처리
    if end_of_sheet_gifts:
        seen_sheet_gifts = set()
        for g, base_pool in end_of_sheet_gifts:
            g_sig = (g.get("order_id"), g.get("product_name"), g.get("option_name"), g.get("reason"))
            if g_sig not in seen_sheet_gifts:
                seen_sheet_gifts.add(g_sig)
                rows.append(make_gift_row(g, base_pool))

    res_df = pd.DataFrame(rows)
    ordered_cols = [c for c in brand_cols if c in res_df.columns]
    meta_cols = [c for c in res_df.columns if c not in brand_cols]
    return res_df[ordered_cols + meta_cols]

def build_daily_summary_sheets(df, shipping_config, event_config=None, custom_item_order=None):
    """
    일자별 수량표: 날짜를 가로 열(Columns)로, 품목(상품명/옵션)을 세로 행(Rows)으로 배치!
    하단에 주문건수, 유료배송, 무료배송, 제주, 이벤트 요약 행을 가로 날짜 열에 맞춰 일체형으로 배치!
    일자별 배송정산표: 주문건별 합산금액을 바탕으로 유료/무료 배송 및 제주추가비 정산
    calc_df: 배송비 구분 및 합산금액이 포함된 전체 발주 데이터프레임
    """
    free_threshold = shipping_config.get("free_shipping_threshold", 50000)
    base_shipping_fee = shipping_config.get("base_shipping_fee", 3000)
    jeju_extra_fee = shipping_config.get("jeju_extra_fee", 3000)

    calc_df = df.copy()
    calc_df["is_jeju"] = calc_df["is_jeju"].apply(
        lambda x: True if (x is True or str(x).strip() in ["True", "1", "제주/도서산간", "제주"]) else False
    )
    calc_df["product_name"] = calc_df["product_name"].fillna("").astype(str)
    calc_df["option_name"] = calc_df["option_name"].fillna("").astype(str)

    # 1. 묶음 식별자 구성 (체크박스로 선택된 필드들만 결합)
    bundle_targets = shipping_config.get("bundle_targets")
    if bundle_targets is None:
        bundle_rule = shipping_config.get("bundle_rule", "")
        if "수취인명 + 연락처 + 주소" in bundle_rule:
            bundle_targets = ["수취인명", "주소지", "연락처"]
        elif "수취인명 + 주소" in bundle_rule:
            bundle_targets = ["수취인명", "주소지"]
        elif "수취인명 + 연락처" in bundle_rule:
            bundle_targets = ["수취인명", "연락처"]
        elif "주문번호 + 수취인명" in bundle_rule:
            bundle_targets = ["주문번호", "수취인명"]
        elif "구매자명 + 연락처" in bundle_rule:
            bundle_targets = ["구매자명", "연락처"]
        elif "주문번호" in bundle_rule:
            bundle_targets = ["주문번호"]
        else:
            bundle_targets = ["수취인명", "주소지", "연락처"]

    b_parts = []
    if "주문번호" in bundle_targets:
        b_parts.append(calc_df["order_id"].fillna("").astype(str))
    if "수취인명" in bundle_targets:
        b_parts.append(calc_df["receiver_name"].fillna("").astype(str))
    if "주소지" in bundle_targets:
        b_parts.append(calc_df["full_address"].fillna("").astype(str))
    if "연락처" in bundle_targets:
        b_parts.append(calc_df["receiver_phone"].fillna("").astype(str))
    if "구매자명" in bundle_targets:
        b_parts.append(calc_df["buyer_name"].fillna("").astype(str))

    if b_parts:
        calc_df["bundle_id"] = b_parts[0]
        for p in b_parts[1:]:
            calc_df["bundle_id"] = calc_df["bundle_id"] + "___" + p
    else:
        calc_df["bundle_id"] = calc_df["order_id"].fillna("").astype(str)

    # 2. 묶음배송 그룹 합산 및 배송비 판정
    order_groups = calc_df.groupby(["order_date", "bundle_id"]).agg({
        "price": "sum",
        "is_jeju": "any",
        "quantity": "sum",
        "order_id": "nunique"
    }).reset_index()

    order_groups["기본배송비"] = order_groups["price"].apply(
        lambda p: 0 if p >= free_threshold else base_shipping_fee
    )
    order_groups["제주추가비"] = order_groups["is_jeju"].apply(
        lambda j: jeju_extra_fee if j else 0
    )
    order_groups["총배송비"] = order_groups["기본배송비"] + order_groups["제주추가비"]

    # 묶음별 정보를 calc_df의 각 행에 매핑 (마스터 Raw DB 상세 확인용 ⭐)
    bundle_price_map = order_groups.set_index(["order_date", "bundle_id"])["price"].to_dict()
    bundle_base_fee_map = order_groups.set_index(["order_date", "bundle_id"])["기본배송비"].to_dict()
    bundle_jeju_fee_map = order_groups.set_index(["order_date", "bundle_id"])["제주추가비"].to_dict()
    bundle_tot_fee_map = order_groups.set_index(["order_date", "bundle_id"])["총배송비"].to_dict()

    def get_bundle_info(row):
        key = (row["order_date"], row["bundle_id"])
        tot_p = int(bundle_price_map.get(key, row.get("price", 0)))
        b_fee = int(bundle_base_fee_map.get(key, 0))
        j_fee = int(bundle_jeju_fee_map.get(key, 0))
        t_fee = int(bundle_tot_fee_map.get(key, 0))
        
        if b_fee == 0 and j_fee == 0:
            status = "🟢 무료배송"
        elif b_fee == 0 and j_fee > 0:
            status = f"🏝️ 무료(제주추가 {j_fee:,}원)"
        elif b_fee > 0 and j_fee == 0:
            status = f"🚚 유료배송 ({b_fee:,}원)"
        else:
            status = f"🚚 유료배송+제주 ({t_fee:,}원)"
            
        return pd.Series([tot_p, b_fee, j_fee, t_fee, status])

    calc_df[["묶음주문합계", "기본배송비", "제주추가비", "부과배송비", "배송비구분"]] = calc_df.apply(get_bundle_info, axis=1)

    daily_shipping_summary = order_groups.groupby("order_date").agg(
        총_배송건수=("bundle_id", "count"),
        총_주문건수=("order_id", "sum"),
        총_구매수량=("quantity", "sum"),
        총_주문금액=("price", "sum"),
        무료배송_건수=("기본배송비", lambda s: (s == 0).sum()),
        유료배송_건수=("기본배송비", lambda s: (s > 0).sum()),
        제주도서_건수=("is_jeju", lambda s: s.sum()),
        기본배송비_합계=("기본배송비", "sum"),
        제주배송비_합계=("제주추가비", "sum"),
        총_배송비_합계=("총배송비", "sum")
    ).reset_index()
    daily_shipping_summary.rename(columns={"order_date": "주문일자"}, inplace=True)

    # 3. 일자별 품목 수량 피벗표 생성
    item_df = calc_df.copy()
    if item_df.empty:
        daily_qty_pivot = pd.DataFrame(columns=["브랜드상품명", "옵션명", "판매단가", "총수량", "총매출액"])
    else:
        item_df["product_name"] = item_df["product_name"].fillna("").astype(str)
        item_df["option_name"] = item_df["option_name"].fillna("").astype(str)
        # 단가 사전 구축
        price_lookup = item_df.groupby(["product_name", "option_name"])["unit_price"].max().to_dict()

        # 날짜별 수량 피벗 (날짜 정렬 보장, dropna=False 로 빈 옵션명 보존)
        qty_pivot = pd.pivot_table(
            item_df,
            index=["product_name", "option_name"],
            columns="order_date",
            values="quantity",
            aggfunc="sum",
            fill_value=0,
            dropna=False
        )
        date_cols = sorted(list(qty_pivot.columns))
        qty_pivot = qty_pivot[date_cols]

        qty_pivot["총수량"] = qty_pivot[date_cols].sum(axis=1)
        qty_pivot.reset_index(inplace=True)
        qty_pivot.rename(columns={"product_name": "브랜드상품명", "option_name": "옵션명"}, inplace=True)

        qty_pivot["판매단가"] = qty_pivot.apply(
            lambda r: price_lookup.get((r["브랜드상품명"], r["옵션명"]), 0), axis=1
        )
        qty_pivot["총매출액"] = qty_pivot["총수량"] * qty_pivot["판매단가"]

        # 3단계 지정 순서 또는 앞 번호순(1, 2, 3...) 우선 정렬 ⭐
        def get_item_order_rank(r):
            bp = str(r["브랜드상품명"]).strip()
            bo = str(r["옵션명"]).strip()
            tot_rev = r.get("총매출액", 0)
            # 1순위: 3단계에서 직접 지정/저장된 순서
            if custom_item_order:
                if (bp, bo) in custom_item_order:
                    return (0, custom_item_order[(bp, bo)], -tot_rev)
                if bp in custom_item_order:
                    return (0, custom_item_order[bp], -tot_rev)
            # 2순위: 품목명 앞의 번호 (1., 2., 3...)
            num = extract_leading_number(bp)
            if num is not None:
                return (1, num, -tot_rev)
            # 3순위: 그 외 일반 품목 / 사은품은 총매출액 높은 순
            return (2, 999999, -tot_rev)

        qty_pivot["_order_rank"] = qty_pivot.apply(get_item_order_rank, axis=1)
        qty_pivot.sort_values(by="_order_rank", inplace=True)
        qty_pivot.drop(columns=["_order_rank"], inplace=True)

        final_cols = ["브랜드상품명", "옵션명", "판매단가"] + date_cols + ["총수량", "총매출액"]
        daily_qty_pivot = qty_pivot[final_cols].copy()

        # 4. [전체 합계] 요약 행
        total_row = {
            "브랜드상품명": "[전체 합계]",
            "옵션명": "-",
            "판매단가": "-"
        }
        for d in date_cols:
            total_row[d] = int(daily_qty_pivot[d].sum())
        total_row["총수량"] = int(daily_qty_pivot["총수량"].sum())
        total_row["총매출액"] = int(daily_qty_pivot["총매출액"].sum())

        # 5. 하단 배송비/이벤트 통합 요약 행 추가 (사용자 요청 Image 2 스타일 일체형 배치 ⭐)
        ship_lookup = daily_shipping_summary.set_index("주문일자").to_dict(orient="index")

        r_orders = {"브랜드상품명": "주문건수", "옵션명": "", "판매단가": ""}
        r_paid = {"브랜드상품명": "유료배송", "옵션명": "", "판매단가": ""}
        r_free = {"브랜드상품명": "무료배송", "옵션명": "", "판매단가": ""}
        r_jeju = {"브랜드상품명": "제주", "옵션명": "", "판매단가": ""}
        r_blank = {"브랜드상품명": "", "옵션명": "", "판매단가": ""}

        tot_orders = 0
        tot_paid = 0
        tot_free = 0
        tot_jeju = 0
        tot_paid_fee = 0

        for d in date_cols:
            d_info = ship_lookup.get(d, {})
            n_order = int(d_info.get("총_배송건수", 0))
            n_paid = int(d_info.get("유료배송_건수", 0))
            n_free = int(d_info.get("무료배송_건수", 0))
            n_jeju = int(d_info.get("제주도서_건수", 0))
            fee_paid = int(d_info.get("기본배송비_합계", 0))

            r_orders[d] = n_order
            r_paid[d] = n_paid
            r_free[d] = n_free
            r_jeju[d] = n_jeju
            r_blank[d] = ""

            tot_orders += n_order
            tot_paid += n_paid
            tot_free += n_free
            tot_jeju += n_jeju
            tot_paid_fee += fee_paid

        r_orders["총수량"] = tot_orders
        r_orders["총매출액"] = ""

        r_paid["총수량"] = tot_paid
        r_paid["총매출액"] = tot_paid_fee

        r_free["총수량"] = tot_free
        r_free["총매출액"] = 0

        r_jeju["총수량"] = tot_jeju
        r_jeju["총매출액"] = 0

        r_blank["총수량"] = ""
        r_blank["총매출액"] = ""

        # 이벤트/사은품별 개별 행 구성 (사용자 요청: [종류] 사은품명 으로 분리 집계 ⭐)
        event_rows = []
        if event_config:
            def clean_gift_label(kind, raw_name):
                clean = re.sub(r'^\[(?:증정|사은품|선착순|금액달성|특정품목|커스텀)\]\s*', '', str(raw_name)).strip()
                return f"[{kind}] {clean}"

            # 1) 선착순 이벤트
            fcfs_c = event_config.get("first_come_first_served", {})
            if fcfs_c.get("enabled"):
                lbl = clean_gift_label("선착순", fcfs_c.get("gift_product_name", "선착순 사은품"))
                r_ev = {"브랜드상품명": lbl, "옵션명": fcfs_c.get("gift_option_name", ""), "판매단가": ""}
                tot_ev = 0
                for d in date_cols:
                    sub_d = calc_df[calc_df["order_date"] == d]
                    cnt = 0
                    for _, r in sub_d.iterrows():
                        for g in r.get("gifts_to_add", []):
                            if "선착순" in str(g.get("reason", "")) or str(g.get("product_name")) == str(fcfs_c.get("gift_product_name")):
                                cnt += g.get("quantity", 1)
                        if cnt == 0:
                            for ev in r.get("applied_events", []):
                                if ev.get("event_type") == "fcfs":
                                    cnt += fcfs_c.get("gift_quantity", 1)
                                    break
                    r_ev[d] = cnt
                    tot_ev += cnt
                r_ev["총수량"] = tot_ev
                r_ev["총매출액"] = ""
                event_rows.append(r_ev)

            # 2) 금액달성 이벤트
            amt_c = event_config.get("min_amount", {})
            if amt_c.get("enabled"):
                lbl = clean_gift_label("금액달성", amt_c.get("gift_product_name", "금액달성 사은품"))
                r_ev = {"브랜드상품명": lbl, "옵션명": amt_c.get("gift_option_name", ""), "판매단가": ""}
                tot_ev = 0
                for d in date_cols:
                    sub_d = calc_df[calc_df["order_date"] == d]
                    cnt = 0
                    for _, r in sub_d.iterrows():
                        for g in r.get("gifts_to_add", []):
                            if "금액달성" in str(g.get("reason", "")) or str(g.get("product_name")) == str(amt_c.get("gift_product_name")):
                                cnt += g.get("quantity", 1)
                        if cnt == 0:
                            for ev in r.get("applied_events", []):
                                if ev.get("event_type") == "amt":
                                    cnt += amt_c.get("gift_quantity", 1)
                                    break
                    r_ev[d] = cnt
                    tot_ev += cnt
                r_ev["총수량"] = tot_ev
                r_ev["총매출액"] = ""
                event_rows.append(r_ev)

            # 3) 특정품목 구매 사은품
            spec_c = event_config.get("special_product", {})
            if spec_c.get("enabled"):
                lbl = clean_gift_label("특정품목", spec_c.get("gift_product_name", "특정품목 증정"))
                r_ev = {"브랜드상품명": lbl, "옵션명": spec_c.get("gift_option_name", ""), "판매단가": ""}
                tot_ev = 0
                for d in date_cols:
                    sub_d = calc_df[calc_df["order_date"] == d]
                    cnt = 0
                    for _, r in sub_d.iterrows():
                        for g in r.get("gifts_to_add", []):
                            if "구매 사은품" in str(g.get("reason", "")) or str(g.get("product_name")) == str(spec_c.get("gift_product_name")):
                                cnt += g.get("quantity", 1)
                        if cnt == 0:
                            for ev in r.get("applied_events", []):
                                if ev.get("event_type") == "spec":
                                    cnt += spec_c.get("gift_quantity", 1)
                                    break
                    r_ev[d] = cnt
                    tot_ev += cnt
                r_ev["총수량"] = tot_ev
                r_ev["총매출액"] = ""
                event_rows.append(r_ev)

            # 4) 커스텀 사은품 목록
            for c_ev in event_config.get("custom_events", []):
                if c_ev.get("enabled"):
                    c_title = c_ev.get("title", "커스텀")
                    lbl = clean_gift_label(c_title, c_ev.get("gift_product_name", "추가 사은품"))
                    r_ev = {"브랜드상품명": lbl, "옵션명": c_ev.get("gift_option_name", ""), "판매단가": ""}
                    tot_ev = 0
                    for d in date_cols:
                        sub_d = calc_df[calc_df["order_date"] == d]
                        cnt = 0
                        for _, r in sub_d.iterrows():
                            for g in r.get("gifts_to_add", []):
                                if c_title in str(g.get("reason", "")) or str(g.get("product_name")) == str(c_ev.get("gift_product_name")):
                                    cnt += g.get("quantity", 1)
                            if cnt == 0:
                                for ev in r.get("applied_events", []):
                                    if ev.get("reason") == c_title or ev.get("event_type") == "custom":
                                        cnt += c_ev.get("gift_quantity", 1)
                                        break
                        r_ev[d] = cnt
                        tot_ev += cnt
                    r_ev["총수량"] = tot_ev
                    r_ev["총매출액"] = ""
                    event_rows.append(r_ev)

        # 등록된 이벤트가 없는 경우 기본 이벤트 행 처리
        if not event_rows:
            r_event = {"브랜드상품명": "이벤트", "옵션명": "", "판매단가": ""}
            tot_event = 0
            for d in date_cols:
                sub_d = calc_df[calc_df["order_date"] == d]
                ev_count = sub_d[sub_d["product_name"].str.contains(r"증정|사은품|<|\[선착순", na=False)]["quantity"].sum()
                ev_tagged = sum(1 for _, r in sub_d.iterrows() if r.get("event_tags") or r.get("applied_events"))
                if ev_count == 0 and ev_tagged > 0:
                    ev_count = ev_tagged
                r_event[d] = int(ev_count)
                tot_event += int(ev_count)
            r_event["총수량"] = tot_event
            r_event["총매출액"] = ""
            event_rows = [r_event]

        summary_rows = [total_row, r_orders, r_paid, r_free, r_jeju, r_blank] + event_rows
        daily_qty_pivot = pd.concat([daily_qty_pivot, pd.DataFrame(summary_rows)], ignore_index=True)

    return daily_qty_pivot, daily_shipping_summary, calc_df

def export_styled_excel(brand_df, daily_qty_df, daily_shipping_df, master_raw_df, color_event="#FFF2CC", color_gift="#E2EFDA"):
    # _row_highlight 및 _row_color 메타데이터 추출 후 엑셀 작성 시 제외
    brand_highlights = []
    brand_colors = []
    brand_df_clean = brand_df.copy()

    if "_row_highlight" in brand_df_clean.columns:
        brand_highlights = brand_df_clean["_row_highlight"].tolist()
        brand_df_clean = brand_df_clean.drop(columns=["_row_highlight"])

    if "_row_color" in brand_df_clean.columns:
        brand_colors = brand_df_clean["_row_color"].tolist()
        brand_df_clean = brand_df_clean.drop(columns=["_row_color"])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        brand_df_clean.to_excel(writer, sheet_name="브랜드사_발주서", index=False)
        daily_qty_df.to_excel(writer, sheet_name="일자별_품목수량표", index=False)
        master_raw_df.to_excel(writer, sheet_name="셀러통합_판매DB", index=False)

    output.seek(0)
    wb = openpyxl.load_workbook(output)

    header_fill = PatternFill(start_color="1E3D59", end_color="1E3D59", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="맑은 고딕", size=9)
    
    clean_ev_hex = (color_event or "#FFF2CC").replace("#", "").upper()
    clean_gift_hex = (color_gift or "#E2EFDA").replace("#", "").upper()
    highlight_fill = PatternFill(start_color=clean_ev_hex, end_color=clean_ev_hex, fill_type="solid")
    gift_fill = PatternFill(start_color=clean_gift_hex, end_color=clean_gift_hex, fill_type="solid")
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
            is_total_row = False
            is_summary_row = False
            is_event_summary_row = False
            is_blank_row = False
            custom_fill = None

            if sheetname == "브랜드사_발주서":
                data_idx = row_idx - 2
                if 0 <= data_idx < len(brand_colors) and brand_colors[data_idx]:
                    c_hex = str(brand_colors[data_idx]).replace("#", "").upper().strip()
                    if len(c_hex) == 6:
                        custom_fill = PatternFill(start_color=c_hex, end_color=c_hex, fill_type="solid")

                # 1순위: 메타데이터 _row_highlight 기반 판정
                is_explicitly_none = False
                if 0 <= data_idx < len(brand_highlights):
                    hl_val = brand_highlights[data_idx]
                    if hl_val == "gift":
                        is_gift_row = True
                    elif hl_val == "event":
                        is_event_tagged = True
                    elif hl_val == "none":
                        is_explicitly_none = True

                # 2순위: 텍스트 기반 판정 (사은품 행 등 보조) - 단 색인 제외(none) 행은 매칭 생략!
                if not is_gift_row and not is_event_tagged and not is_explicitly_none:
                    row_str = " ".join(str(ws.cell(row=row_idx, column=c).value or "") for c in range(1, ws.max_column + 1))
                    if "[사은품]" in row_str:
                        is_gift_row = True
                    elif any(kw in row_str for kw in ["[선착순", "원이상", "구매]"]):
                        is_event_tagged = True
            elif sheetname == "일자별_품목수량표":
                val_first = str(ws.cell(row=row_idx, column=1).value or "").strip()
                if "[전체 합계]" in val_first:
                    is_total_row = True
                    ws.cell(row=row_idx, column=3).value = "-"
                elif val_first in ["주문건수", "유료배송", "무료배송", "제주"]:
                    is_summary_row = True
                elif val_first == "이벤트" or (val_first.startswith("[") and val_first != "[전체 합계]"):
                    is_event_summary_row = True
                elif val_first == "":
                    is_blank_row = True

            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

                if is_blank_row:
                    cell.fill = PatternFill(fill_type=None)
                    cell.border = Border()
                elif is_total_row:
                    cell.fill = PatternFill(start_color="EBF3FA", end_color="EBF3FA", fill_type="solid")
                    cell.font = Font(name="맑은 고딕", size=9, bold=True, color="1E3D59")
                elif is_summary_row:
                    cell.fill = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid")
                    cell.font = Font(name="맑은 고딕", size=9, bold=True, color="2F3437")
                elif is_event_summary_row:
                    cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                    cell.font = Font(name="맑은 고딕", size=9, bold=True, color="B78103")
                elif custom_fill is not None:
                    cell.fill = custom_fill
                elif is_gift_row:
                    cell.fill = gift_fill
                elif is_event_tagged:
                    cell.fill = highlight_fill

                if not is_blank_row and isinstance(cell.value, (int, float)) and cell.value >= 1000:
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

# =========================================================================
# 3. Streamlit 대시보드 UI
# =========================================================================
brand_templates = load_brand_templates()
seller_mappings = load_seller_mappings()

# ----------------- 사이드바 (Notion Stepper & 진행 상태) -----------------
with st.sidebar:
    st.markdown("""
    <div style="padding: 10px 0 16px 0; border-bottom: 1px solid #ECECEE; margin-bottom: 12px; text-align: left;">
        <div style="font-size: 24px; font-weight: 800; color: #1E3D59; display: flex; align-items: center; gap: 8px; letter-spacing: -0.5px;">
            <span>📑</span> Order Bridge Pro
        </div>
        <div style="font-size: 13px; font-weight: 600; color: #5A6065; margin-top: 5px; text-align: left;">
            발주서 변환 워크스페이스
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 1. 진행 현황 Stepper
    active_brand_curr = st.session_state.get("active_brand", "호무로")
    brand_cols_curr = st.session_state.get("active_brand_cols", [])
    has_seller_df = "current_seller_df" in st.session_state
    seller_file_curr = st.session_state.get("current_seller_file", "")
    active_pm_curr = st.session_state.get("active_product_mapping", {})
    event_skip = st.session_state.get("event_skip_mode", False)
    event_confirmed = st.session_state.get("event_config_confirmed", False)
    ev_cfg_curr = st.session_state.get("event_config", {})
    any_event_on = any(ev_cfg_curr.get(k, {}).get("enabled", False) for k in ["first_come_first_served", "min_amount", "special_product"])

    # 매칭 상태 텍스트 계산
    unmap_badge_text = "⚪ 대기"
    if has_seller_df and st.session_state.current_seller_df is not None:
        try:
            curr_df_clean = sanitize_dataframe_columns(st.session_state.current_seller_df)
            st.session_state.current_seller_df = curr_df_clean
            
            # 위젯 상태에서 최신 선택 매핑 읽기
            active_col_map = {}
            for f_k in STANDARD_FIELDS:
                if f"sel_col_{f_k}" in st.session_state:
                    active_col_map[f_k] = st.session_state[f"sel_col_{f_k}"]
                elif "user_col_map" in st.session_state and f_k in st.session_state.user_col_map:
                    active_col_map[f_k] = st.session_state.user_col_map[f_k]
                    
            raw_items_temp = extract_unique_products(curr_df_clean, active_col_map)
            if raw_items_temp:
                unmapped_items_temp = [item for item in raw_items_temp if f"{item[0]}||{item[1]}" not in active_pm_curr and f"{item[0]}||" not in active_pm_curr]
                is_keep_all = (raw_items_temp and all(active_pm_curr.get(f"{it[0]}||{it[1]}", {}).get("keep_original") for it in raw_items_temp))
                if st.session_state.get("step3_keep_original_mode") or is_keep_all:
                    unmap_badge_text = "🔵 변경없음"
                elif len(unmapped_items_temp) == 0:
                    unmap_badge_text = f"🟢 완료 ({len(raw_items_temp)}건)"
                else:
                    unmap_badge_text = f"🟡 미매칭 {len(unmapped_items_temp)}건"
        except Exception:
            unmap_badge_text = "🟡 확인중"

    brand_badge_text = f"🟢 {active_brand_curr}" if brand_cols_curr else "⚪ 미선택"
    
    active_seller_display = st.session_state.get("active_seller_name") or seller_file_curr
    if has_seller_df:
        seller_badge_text = f"🟢 {active_seller_display}"
    else:
        seller_badge_text = "⚪ 대기"
    
    if event_skip:
        event_badge_text = "⏩ 행사없음"
    elif any_event_on:
        event_badge_text = "🟢 설정완료"
    else:
        event_badge_text = "⚪ 기본 OFF"

    step5_badge_text = "🟢 준비완료" if has_seller_df else "⚪ 대기"

    if "active_step_idx" not in st.session_state:
        st.session_state["active_step_idx"] = 0

    cur_step = st.session_state.get("active_step_idx", 0)

    # 클릭 가능한 인터랙티브 Stepper 버튼 바
    st.markdown("""
    <div style="font-size: 13px; font-weight: 700; color: #787774; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">
        📌 작업 진행 단계 (클릭하여 이동 👆)
    </div>
    """, unsafe_allow_html=True)

    steps_meta = [
        ("1. 브랜드 양식", brand_badge_text),
        ("2. 셀러 발주서", seller_badge_text),
        ("3. 품목·단가 매칭", unmap_badge_text),
        ("4. 이벤트 & 사은품", event_badge_text),
        ("5. 최종 발주서 생성", step5_badge_text)
    ]

    for s_idx, (s_label, s_badge) in enumerate(steps_meta):
        is_cur = (cur_step == s_idx)
        btn_label = f"{'👉 ' if is_cur else '   '}{s_label}\n     {s_badge}"
        if st.button(btn_label, key=f"sidebar_step_nav_btn_{s_idx}", use_container_width=True, type="primary" if is_cur else "secondary"):
            st.session_state["active_step_idx"] = s_idx
            st.rerun()

    st.write(" ")

    # 2. 브랜드 배송 정책 요약 안내 카드 (공용 입력 대신 활성 브랜드 정책 표시 ⭐)
    active_b = st.session_state.get("active_brand", "호무로")
    active_tmpl_k = st.session_state.get("active_brand_template_key", active_b)
    active_b_data = brand_templates.get(active_tmpl_k) or brand_templates.get(active_b, {})
    curr_ship_cfg = st.session_state.get("active_shipping_config") or active_b_data.get("shipping_config", {
        "bundle_targets": ["수취인명", "주소지", "연락처"],
        "free_shipping_threshold": 50000,
        "base_shipping_fee": 3000,
        "jeju_extra_fee": 3000
    })

    with st.expander(f"🚚 배송비 & 묶음배송 정책 ({active_b})", expanded=False):
        b_targets_list = curr_ship_cfg.get("bundle_targets")
        if b_targets_list is None:
            b_targets_list = ["수취인명", "주소지", "연락처"]
        b_targets_str = " + ".join(b_targets_list) if b_targets_list else "묶음배송 미적용 (주문번호별 개별 발송)"
        st.markdown(f"**📦 묶음배송 기준**: `{b_targets_str}`")
        st.markdown(f"**🟢 무료배송 기준**: `{curr_ship_cfg.get('free_shipping_threshold', 50000):,}원 이상`")
        st.markdown(f"**🚚 기본 배송비**: `{curr_ship_cfg.get('base_shipping_fee', 3000):,}원`")
        st.markdown(f"**🏝️ 제주 추가비**: `+{curr_ship_cfg.get('jeju_extra_fee', 3000):,}원`")
        st.caption("💡 브랜드마다 배송 정책이 다르므로, 변경은 **1단계 브랜드 양식**에서 가능합니다.")

    # 3. 통합 백업 및 보관함 expander
    with st.expander("💾 통합 백업 및 데이터 보관함", expanded=False):
        st.caption("현재 등록된 모든 브랜드사 양식과 셀러 매핑 정보를 JSON으로 일괄 백업하거나 복원합니다.")
        full_json_str = export_full_backup_json()
        now_dt = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button(
            label="📥 전체 데이터 통합 백업 다운로드 (.json)",
            data=full_json_str,
            file_name=f"OrderBridge_통합백업_{now_dt}.json",
            mime="application/json",
            use_container_width=True
        )
        st.write("")
        uploaded_backup = st.file_uploader("📤 백업 파일 복원 (.json)", type=["json", "JSON"], key="uploader_full_backup")
        if uploaded_backup is not None:
            if st.button("🔄 백업 파일 적용 및 덮어쓰기", type="primary", use_container_width=True):
                content = uploaded_backup.read().decode("utf-8")
                ok, msg = import_full_backup_json(content)
                if ok:
                    st.success("🎉 데이터가 성공적으로 복원되었습니다!")
                    st.rerun()
                else:
                    st.error(msg)
        st.caption(f"📂 저장 경로: `{CURRENT_DIR}`")

    # 최하단 푸터 정보 (가운데 정렬, 투명도 50%)
    st.markdown("""
    <div style="margin-top: 28px; padding-top: 14px; border-top: 1px solid #ECECEE; text-align: center; opacity: 0.5; font-size: 12px; color: #37352F; font-weight: 500;">
        제작자 : 태민 / 버전 1.0
    </div>
    """, unsafe_allow_html=True)

shipping_config = curr_ship_cfg

# 기본 이벤트 설정 초기화 (3가지 핵심 프로모션 + 커스텀 사은품 + 색인 색상표 ⭐)
default_event_config = {
    "color_event": "#FFF2CC",
    "color_gift": "#E2EFDA",
    "first_come_first_served": {
        "enabled": False,
        "limit": 3,
        "gift_product_name": "[증정] 선착순 미니어처 15ml",
        "gift_option_name": "증정용",
        "gift_quantity": 1,
        "target_rule": ["수취인명", "연락처"],
        "tag_display_rule": "텍스트 표기 없음 (화면 및 엑셀 배경색 하이라이트 색인만 처리 ⭐)",
        "gift_position": "해당 주문(합포장) 가장 아래에 사은품 행 추가 (추천 ⭐)",
        "action_type": "add_row",
        "write_col_enabled": False,
        "write_target_col": "N"
    },
    "min_amount": {
        "enabled": False,
        "min_amount": 50000,
        "fcfs_enabled": False,
        "fcfs_limit": 50,
        "gift_product_name": "[증정] 금액달성 사은품",
        "gift_option_name": "1세트",
        "gift_quantity": 1,
        "target_rule": ["주문번호"],
        "tag_display_rule": "텍스트 표기 없음 (화면 및 엑셀 배경색 하이라이트 색인만 처리 ⭐)",
        "gift_position": "해당 주문(합포장) 가장 아래에 사은품 행 추가 (추천 ⭐)",
        "write_col_enabled": False,
        "write_target_col": "N"
    },
    "special_product": {
        "enabled": False,
        "keyword": "",
        "gift_product_name": "[사은품] 특정품목 증정",
        "gift_option_name": "단품",
        "gift_quantity": 1,
        "target_rule": ["주문번호"],
        "tag_display_rule": "텍스트 표기 없음 (화면 및 엑셀 배경색 하이라이트 색인만 처리 ⭐)",
        "gift_position": "해당 주문(합포장) 가장 아래에 사은품 행 추가 (추천 ⭐)",
        "write_col_enabled": False,
        "write_target_col": "N"
    },
    "custom_events": []
}

if "event_config" not in st.session_state:
    st.session_state["event_config"] = default_event_config

event_config = st.session_state.get("event_config", default_event_config)

# ----------------- 구글 스타일 팔레트 + 직접 색상 선택 헬퍼 -----------------
EVENT_PALETTE_PRESETS = [
    ("연노랑 💛", "#FFF2CC"),
    ("연초록 💚", "#E2EFDA"),
    ("연파랑 💙", "#D9E1F2"),
    ("연살구 🧡", "#FCE4D6"),
    ("연분홍 💖", "#F8CECC"),
    ("연보라 💜", "#E1D5E7"),
    ("하늘색 🩵", "#D0F0FD"),
    ("연회색 🩶", "#EDEDED"),
]
NAME_TO_HEX = {name: h.upper() for name, h in EVENT_PALETTE_PRESETS}
HEX_TO_NAME = {h.upper(): name for name, h in EVENT_PALETTE_PRESETS}

def on_ev_palette_pill_change(key_prefix):
    chosen_pill = st.session_state.get(f"{key_prefix}_pill")
    if chosen_pill and chosen_pill in NAME_TO_HEX:
        st.session_state[f"{key_prefix}_color_picker"] = NAME_TO_HEX[chosen_pill]

def on_ev_color_picker_change(key_prefix):
    picked_hex = str(st.session_state.get(f"{key_prefix}_color_picker", "")).upper()
    st.session_state[f"{key_prefix}_pill"] = HEX_TO_NAME.get(picked_hex, None)

ACTION_RULE_OPTIONS: list[str] = [
    "해당 주문(합포장) 가장 아래에 사은품 행 추가 (추천 ⭐)",
    "발주서 최하단(맨 끝)에 사은품만 모아서 일괄 추가",
    "별도 행 추가 안 함 (배경색 하이라이트 색인만 ⭐)",
    "별도 행 추가 안 함 (배송메세지에 텍스트 표기)"
]

def render_google_style_color_selector(label: str, default_hex: str, key_prefix: str, default_no_color: bool = False) -> tuple[str, bool]:
    """구글 폰트/문서 스타일: 파스텔 추천 팔레트 빠른 선택 + 직접 선택(Color Picker) 결합 위젯"""
    init_hex: str = str(default_hex if default_hex else "#FFF2CC").upper()
    if f"{key_prefix}_color_picker" not in st.session_state:
        st.session_state[f"{key_prefix}_color_picker"] = init_hex
    if f"{key_prefix}_pill" not in st.session_state:
        st.session_state[f"{key_prefix}_pill"] = HEX_TO_NAME.get(init_hex, None)

    st.markdown(f"**🎨 {label} (팔레트 빠른 선택 + 직접 색상 선택)**")
    
    col_pal, col_pick, col_badge, col_chk = st.columns([3.4, 0.9, 0.9, 1.4])
    with col_pal:
        st.pills(
            "추천 팔레트",
            options=list(NAME_TO_HEX.keys()),
            key=f"{key_prefix}_pill",
            on_change=on_ev_palette_pill_change,
            args=(key_prefix,),
            label_visibility="collapsed"
        )
    with col_pick:
        st.color_picker(
            "직접 선택 (+)",
            value=st.session_state.get(f"{key_prefix}_color_picker", init_hex),
            key=f"{key_prefix}_color_picker",
            on_change=on_ev_color_picker_change,
            args=(key_prefix,),
            label_visibility="collapsed"
        )
    with col_badge:
        cur_hex: str = str(st.session_state.get(f"{key_prefix}_color_picker", init_hex)).upper()
        st.markdown(f"""
        <div style="padding: 4px 6px; background: {cur_hex}; border: 1px solid #CBD5E1; border-radius: 6px; text-align: center; font-weight: 700; font-size: 11px; color: #1E293B; line-height: 22px;">
            {cur_hex}
        </div>
        """, unsafe_allow_html=True)
    with col_chk:
        no_color: bool = st.checkbox("🚫 색인(배경색) 안 함", value=default_no_color, key=f"{key_prefix}_no_color")
    
    return cur_hex, no_color

def apply_event_template_to_session(loaded_cfg: dict, tmpl_name: str, default_ev_title: str) -> None:
    """
    이벤트 템플릿(선착순, 금액달성, 특정품목, 커스텀 이벤트)의 설정 데이터를
    Streamlit 세션 상태의 모든 위젯 키에 100% 완벽하게 바인딩하여 복원합니다.
    """
    # 1. 기존 이벤트 관련 세션 키 정리 (이전 템플릿 잔여값 완전 소거)
    all_keys = list(st.session_state.keys())
    for k in all_keys:
        if k.startswith(("ev_", "fcfs_", "amt_", "spec_", "c_title_", "c_on_", "c_cond_", "c_limit_", 
                         "c_min_amt_", "c_kw_", "c_k_", "c_act_rule_", "c_write_col_on_", "c_target_col_", 
                         "c_qty_", "c_pname_", "c_opt_", "c_hl_scope_", "c_del_", "c_ev_")):
            st.session_state.pop(k, None)

    cfg: dict = loaded_cfg if isinstance(loaded_cfg, dict) else {}

    # 2. 선착순 이벤트 (first_come_first_served)
    fcfs: dict = cfg.get("first_come_first_served", {})
    st.session_state["ev_fcfs_on"] = bool(fcfs.get("enabled", False))
    st.session_state["ev_fcfs_limit"] = int(fcfs.get("limit", 3))
    fcfs_targets = fcfs.get("target_rule", ["수취인명", "연락처"])
    if isinstance(fcfs_targets, str):
        fcfs_targets = [fcfs_targets]
    st.session_state["fcfs_chk_oid"] = any("주문번호" in str(x) for x in fcfs_targets)
    st.session_state["fcfs_chk_name"] = any(k in str(x) for x in fcfs_targets for k in ["수취인", "수령인", "구매자", "1인"])
    st.session_state["fcfs_chk_phone"] = any(k in str(x) for x in fcfs_targets for k in ["연락처", "전화", "1인"])
    st.session_state["fcfs_chk_addr"] = any("주소" in str(x) for x in fcfs_targets)
    
    fcfs_act = fcfs.get("gift_position", ACTION_RULE_OPTIONS[0])
    for opt in ACTION_RULE_OPTIONS:
        if opt in fcfs_act or fcfs_act in opt:
            fcfs_act = opt
            break
    st.session_state["ev_fcfs_act_rule"] = fcfs_act
    st.session_state["ev_fcfs_write_col_on"] = bool(fcfs.get("write_col_enabled", False))
    st.session_state["ev_fcfs_target_col"] = str(fcfs.get("write_target_col", "N"))
    st.session_state["ev_fcfs_qty"] = int(fcfs.get("gift_quantity", 1))
    st.session_state["ev_fcfs_pname"] = str(fcfs.get("gift_product_name", "[증정] 선착순 미니어처 15ml"))
    st.session_state["ev_fcfs_opt"] = str(fcfs.get("gift_option_name", "증정용"))
    st.session_state["ev_fcfs_hl_scope"] = "대상자들에 색인 (본품+사은품)" if fcfs.get("highlight_scope", "both") == "both" else "이벤트에만 색인 (사은품만)"
    fcfs_hex = str(fcfs.get("color", "#FFF2CC")).upper()
    st.session_state["ev_fcfs_color_picker"] = fcfs_hex
    st.session_state["ev_fcfs_pill"] = HEX_TO_NAME.get(fcfs_hex, None)
    st.session_state["ev_fcfs_no_color"] = bool(fcfs.get("no_color", False))

    # 3. 구매금액별 이벤트 (min_amount)
    amt: dict = cfg.get("min_amount", {})
    st.session_state["ev_amt_on"] = bool(amt.get("enabled", False))
    st.session_state["ev_amt_limit"] = int(amt.get("min_amount", 50000))
    amt_targets = amt.get("target_rule", ["주문번호"])
    if isinstance(amt_targets, str):
        amt_targets = [amt_targets]
    st.session_state["amt_chk_oid"] = any("주문번호" in str(x) for x in amt_targets)
    st.session_state["amt_chk_name"] = any(k in str(x) for x in amt_targets for k in ["수취인", "수령인", "구매자"])
    st.session_state["amt_chk_phone"] = any(k in str(x) for x in amt_targets for k in ["연락처", "전화"])
    st.session_state["amt_chk_addr"] = any("주소" in str(x) for x in amt_targets)
    st.session_state["amt_chk_multiplier"] = any("배수" in str(x) for x in amt_targets)
    st.session_state["ev_amt_fcfs_on"] = bool(amt.get("fcfs_enabled", False))
    st.session_state["ev_amt_fcfs_limit"] = int(amt.get("fcfs_limit", 30))
    amt_act = amt.get("gift_position", ACTION_RULE_OPTIONS[0])
    for opt in ACTION_RULE_OPTIONS:
        if opt in amt_act or amt_act in opt:
            amt_act = opt
            break
    st.session_state["ev_amt_act_rule"] = amt_act
    st.session_state["ev_amt_write_col_on"] = bool(amt.get("write_col_enabled", False))
    st.session_state["ev_amt_target_col"] = str(amt.get("write_target_col", "N"))
    st.session_state["ev_amt_qty"] = int(amt.get("gift_quantity", 1))
    st.session_state["ev_amt_pname"] = str(amt.get("gift_product_name", "[증정] 금액달성 사은품"))
    st.session_state["ev_amt_opt"] = str(amt.get("gift_option_name", "1세트"))
    st.session_state["ev_amt_hl_scope"] = "대상자들에 색인 (본품+사은품)" if amt.get("highlight_scope", "both") == "both" else "이벤트에만 색인 (사은품만)"
    amt_hex = str(amt.get("color", "#FFF2CC")).upper()
    st.session_state["ev_amt_color_picker"] = amt_hex
    st.session_state["ev_amt_pill"] = HEX_TO_NAME.get(amt_hex, None)
    st.session_state["ev_amt_no_color"] = bool(amt.get("no_color", False))

    # 4. 특정 품목 구매 이벤트 (special_product)
    spec: dict = cfg.get("special_product", {})
    st.session_state["ev_spec_on"] = bool(spec.get("enabled", False))
    st.session_state["ev_spec_kw"] = str(spec.get("keyword", ""))
    spec_targets = spec.get("target_rule", ["주문번호"])
    if isinstance(spec_targets, str):
        spec_targets = [spec_targets]
    st.session_state["spec_chk_oid"] = any("주문번호" in str(x) for x in spec_targets)
    st.session_state["spec_chk_name"] = any(k in str(x) for x in spec_targets for k in ["수취인", "수령인", "구매자"])
    st.session_state["spec_chk_phone"] = any(k in str(x) for x in spec_targets for k in ["연락처", "전화"])
    st.session_state["spec_chk_addr"] = any("주소" in str(x) for x in spec_targets)
    st.session_state["spec_chk_ratio"] = any("1:1" in str(x) for x in spec_targets)
    spec_act = spec.get("gift_position", ACTION_RULE_OPTIONS[0])
    for opt in ACTION_RULE_OPTIONS:
        if opt in spec_act or spec_act in opt:
            spec_act = opt
            break
    st.session_state["ev_spec_act_rule"] = spec_act
    st.session_state["ev_spec_write_col_on"] = bool(spec.get("write_col_enabled", False))
    st.session_state["ev_spec_target_col"] = str(spec.get("write_target_col", "N"))
    st.session_state["ev_spec_qty"] = int(spec.get("gift_quantity", 1))
    st.session_state["ev_spec_pname"] = str(spec.get("gift_product_name", "[사은품] 특정품목 증정"))
    st.session_state["ev_spec_opt"] = str(spec.get("gift_option_name", "단품"))
    st.session_state["ev_spec_hl_scope"] = "대상자들에 색인 (본품+사은품)" if spec.get("highlight_scope", "both") == "both" else "이벤트에만 색인 (사은품만)"
    spec_hex = str(spec.get("color", "#E2EFDA")).upper()
    st.session_state["ev_spec_color_picker"] = spec_hex
    st.session_state["ev_spec_pill"] = HEX_TO_NAME.get(spec_hex, None)
    st.session_state["ev_spec_no_color"] = bool(spec.get("no_color", False))

    # 5. 커스텀 이벤트 (custom_events)
    custom_events: list = cfg.get("custom_events", [])
    st.session_state["custom_events_list"] = custom_events
    for c_i, c_item in enumerate(custom_events):
        st.session_state[f"c_title_{c_i}"] = str(c_item.get("title", f"사은품 이벤트 {c_i+1}"))
        st.session_state[f"c_on_{c_i}"] = bool(c_item.get("enabled", True))
        c_cond_types = ["전체 주문 증정", "선착순 주문 증정", "구매금액 달성 증정", "특정 품목/키워드 증정"]
        curr_cond = c_item.get("condition_type", "전체 주문 증정")
        for ct in c_cond_types:
            if ct in curr_cond or curr_cond in ct:
                curr_cond = ct
                break
        st.session_state[f"c_cond_{c_i}"] = curr_cond
        st.session_state[f"c_limit_{c_i}"] = int(c_item.get("limit", 10))
        st.session_state[f"c_min_amt_{c_i}"] = int(c_item.get("min_amount", 30000))
        st.session_state[f"c_kw_{c_i}"] = str(c_item.get("keyword", ""))
        c_targets = c_item.get("target_rule", ["수취인명", "연락처"])
        if isinstance(c_targets, str):
            c_targets = [c_targets]
        st.session_state[f"c_k_oid_{c_i}"] = any("주문번호" in str(x) for x in c_targets)
        st.session_state[f"c_k_name_{c_i}"] = any(k in str(x) for x in c_targets for k in ["수취인", "수령인", "구매자"])
        st.session_state[f"c_k_phone_{c_i}"] = any(k in str(x) for x in c_targets for k in ["연락처", "전화"])
        st.session_state[f"c_k_addr_{c_i}"] = any("주소" in str(x) for x in c_targets)
        c_act = c_item.get("gift_position", ACTION_RULE_OPTIONS[0])
        for opt in ACTION_RULE_OPTIONS:
            if opt in c_act or c_act in opt:
                c_act = opt
                break
        st.session_state[f"c_act_rule_{c_i}"] = c_act
        st.session_state[f"c_write_col_on_{c_i}"] = bool(c_item.get("write_col_enabled", False))
        st.session_state[f"c_target_col_{c_i}"] = str(c_item.get("write_target_col", "N"))
        st.session_state[f"c_qty_{c_i}"] = int(c_item.get("gift_quantity", 1))
        st.session_state[f"c_pname_{c_i}"] = str(c_item.get("gift_product_name", "[증정] 추가 사은품"))
        st.session_state[f"c_opt_{c_i}"] = str(c_item.get("gift_option_name", "단품"))
        st.session_state[f"c_hl_scope_{c_i}"] = "대상자들에 색인 (본품+사은품)" if c_item.get("highlight_scope", "both") == "both" else "이벤트에만 색인 (사은품만)"
        c_hex = str(c_item.get("color", "#FFF2CC")).upper()
        st.session_state[f"c_ev_{c_i}_color_picker"] = c_hex
        st.session_state[f"c_ev_{c_i}_pill"] = HEX_TO_NAME.get(c_hex, None)
        st.session_state[f"c_ev_{c_i}_no_color"] = bool(c_item.get("no_color", False))

    # 6. 전역 상태 갱신
    st.session_state["event_config"] = cfg
    st.session_state["input_event_tmpl_name_widget"] = tmpl_name if tmpl_name != "+ 새 이벤트 템플릿 설정 (직접 구성)" else default_ev_title
    st.session_state["last_ev_sync_context"] = f"ev_tmpl___{tmpl_name}"

# ----------------- 메인 영역 (큼직한 노션 상단 탭 네비게이터) -----------------
st.markdown("""
<div class="notion-header">
    <div class="notion-title">
        <span>📦</span> Order Bridge Pro — 발주서 자동 변환 시스템
    </div>
    <div class="notion-subtitle">
        1단계 브랜드사 양식 설정 ➔ 2단계 셀러 발주서 업로드 ➔ 3단계 품목/단가 매칭 ➔ 4단계 이벤트 & 사은품 커스텀 ➔ 5단계 최종 발주서 및 일자별 품목수량표
    </div>
</div>
""", unsafe_allow_html=True)

STEP_NAV_OPTIONS = [
    "1️⃣ Step 1. 브랜드사 양식 관리",
    "2️⃣ Step 2. 셀러 발주서 & 헤더 매칭",
    "3️⃣ Step 3. 품목 매칭 & 본사 판매단가",
    "4️⃣ Step 4. 이벤트 & 사은품 맞춤 설정",
    "5️⃣ Step 5. 최종 발주서 & 정산표 다운로드"
]

cur_step_idx = st.session_state.get("active_step_idx", 0)
if cur_step_idx >= len(STEP_NAV_OPTIONS):
    cur_step_idx = 0

selected_step_radio = st.radio(
    "단계 선택",
    options=STEP_NAV_OPTIONS,
    index=cur_step_idx,
    horizontal=True,
    label_visibility="collapsed",
    key=f"top_step_nav_radio_{cur_step_idx}"
)

new_step_idx = STEP_NAV_OPTIONS.index(selected_step_radio)
if new_step_idx != cur_step_idx:
    st.session_state["active_step_idx"] = new_step_idx
    st.rerun()

# =========================================================================
# STEP 1: 브랜드사 양식(템플릿) 설정
# =========================================================================
if cur_step_idx == 0:
    st.markdown("""
    <div class="notion-step-banner">1️⃣ 1단계. 브랜드사 발주 양식 관리</div>
    <div class="notion-step-desc-big">납품할 브랜드사의 발주서 컬럼 순서 및 서식을 설정하세요. 엑셀 파일을 업로드하면 헤더 컬럼을 1초 만에 자동 추출합니다.</div>
    """, unsafe_allow_html=True)
    
    brand_keys = list(brand_templates.keys())
    
    # ---------------- 상단: 기존 템플릿 선택 및 관리 ----------------
    b_col1, b_col2 = st.columns([2, 1])
    with b_col1:
        b_options = brand_keys + ["+ 새 브랜드사 템플릿 등록"]
        active_b_curr = st.session_state.get("active_brand")
        if "pending_brand_choice" in st.session_state:
            curr_b_choice = st.session_state.pop("pending_brand_choice")
        elif active_b_curr and active_b_curr in b_options:
            curr_b_choice = active_b_curr
        else:
            curr_b_choice = st.session_state.get("sb_brand_key_selector", b_options[0] if b_options else "+ 새 브랜드사 템플릿 등록")
        
        if curr_b_choice not in b_options:
            curr_b_choice = b_options[0] if b_options else "+ 새 브랜드사 템플릿 등록"
        def_b_idx = b_options.index(curr_b_choice)

        selected_brand_key = st.selectbox(
            "적용할 브랜드사 템플릿 선택",
            b_options,
            index=def_b_idx,
            key="sb_brand_key_selector"
        )
        if selected_brand_key == "+ 새 브랜드사 템플릿 등록":
            new_tmpl_name_val = st.text_input(
                "새 브랜드사 템플릿 이름 입력",
                value=st.session_state.get("input_new_brand_tmpl", ""),
                placeholder="예: 호무로, 리오마레, 템포 등",
                key="input_new_brand_tmpl"
            ).strip()
            target_brand_name = new_tmpl_name_val
        else:
            target_brand_name = selected_brand_key
    with b_col2:
        if selected_brand_key != "+ 새 브랜드사 템플릿 등록":
            st.write(" ")
            st.write(" ")
            if st.button("🗑️ 이 브랜드사 템플릿 삭제", key="del_brand_tmpl_btn", type="secondary"):
                if selected_brand_key in brand_templates:
                    del brand_templates[selected_brand_key]
                    save_brand_templates(brand_templates)
                    st.warning(f"'{selected_brand_key}' 브랜드사 템플릿이 삭제되었습니다.")
                    st.session_state["last_brand_sync_context"] = None
                    st.rerun()

    if selected_brand_key != "+ 새 브랜드사 템플릿 등록":
        tmpl_entry = brand_templates.get(selected_brand_key, {})
        default_template_name = tmpl_entry.get("template_name", selected_brand_key)
        default_brand_name = tmpl_entry.get("brand_name", selected_brand_key)
        init_cols = tmpl_entry.get("columns", [])
        saved_ship = tmpl_entry.get("shipping_config", {})
    else:
        default_template_name = ""
        default_brand_name = ""
        init_cols = []  # 신규 등록 시 공란으로 시작!
        saved_ship = {}

    txt_area_key = f"header_text_area_{selected_brand_key}"

    # 템플릿 선택 및 전환 시 헤더를 즉시 동기화 ⭐
    sync_brand_context = f"brand_tmpl___{selected_brand_key}"
    if st.session_state.get("last_brand_sync_context") != sync_brand_context:
        if selected_brand_key != "+ 새 브랜드사 템플릿 등록":
            st.session_state[txt_area_key] = ", ".join(init_cols) if init_cols else ""
            st.session_state["active_brand_template_key"] = selected_brand_key
            st.session_state["active_brand"] = selected_brand_key
            st.session_state["active_brand_cols"] = init_cols
            st.session_state["active_shipping_config"] = saved_ship

            # 💡 저장된 브랜드 템플릿의 묶음배송 판정 기준 및 배송비 설정 복원
            b_targets = saved_ship.get("bundle_targets")
            if b_targets is None:
                b_targets = ["수취인명", "주소지", "연락처"]
            st.session_state[f"ship_cb_order_{selected_brand_key}"] = ("주문번호" in b_targets)
            st.session_state[f"ship_cb_name_{selected_brand_key}"] = ("수취인명" in b_targets)
            st.session_state[f"ship_cb_addr_{selected_brand_key}"] = ("주소지" in b_targets)
            st.session_state[f"ship_cb_phone_{selected_brand_key}"] = ("연락처" in b_targets)
            st.session_state[f"ship_free_{selected_brand_key}"] = int(saved_ship.get("free_shipping_threshold", 50000))
            st.session_state[f"ship_base_{selected_brand_key}"] = int(saved_ship.get("base_shipping_fee", 3000))
            st.session_state[f"ship_jeju_{selected_brand_key}"] = int(saved_ship.get("jeju_extra_fee", 3000))
        else:
            st.session_state[txt_area_key] = ""
            st.session_state["active_brand_cols"] = []

            # 💡 새로 작업할 때만 기본값(수취인명, 주소지, 연락처 / 50000원 / 3000원 / 3000원)으로 초기화!
            st.session_state[f"ship_cb_order_{selected_brand_key}"] = False
            st.session_state[f"ship_cb_name_{selected_brand_key}"] = True
            st.session_state[f"ship_cb_addr_{selected_brand_key}"] = True
            st.session_state[f"ship_cb_phone_{selected_brand_key}"] = True
            st.session_state[f"ship_free_{selected_brand_key}"] = 50000
            st.session_state[f"ship_base_{selected_brand_key}"] = 3000
            st.session_state[f"ship_jeju_{selected_brand_key}"] = 3000

        st.session_state.pop("brand_extracted_columns", None)
        st.session_state["last_brand_sync_context"] = sync_brand_context
        st.rerun()

    saved_targets = saved_ship.get("bundle_targets")
    if saved_targets is None:
        saved_targets = ["수취인명", "주소지", "연락처"]
    saved_free = saved_ship.get("free_shipping_threshold", 50000)
    saved_base = saved_ship.get("base_shipping_fee", 3000)
    saved_jeju = saved_ship.get("jeju_extra_fee", 3000)

    st.markdown("---")
    st.markdown("#### 📂 방법 A: 브랜드사 발주서 엑셀 파일을 넣어 자동으로 헤더 가져오기 (가장 추천 ⭐)")
    st.caption("브랜드사에서 준 빈 양식이나 기존 발주서 엑셀을 업로드하면 첫 행의 컬럼들을 1초 만에 자동으로 추출합니다.")
    
    brand_excel_file = st.file_uploader(
        "브랜드사 엑셀 파일 업로드 (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv", "XLSX", "XLS", "CSV"],
        key=f"brand_sample_uploader_{selected_brand_key}"
    )

    # 엑셀 파일 업로드 시 컬럼 및 브랜드명 자동 추출
    if brand_excel_file:
        try:
            df_brand_sample = read_uploaded_tabular_file(brand_excel_file)
            ext_cols = [str(c).strip() for c in df_brand_sample.columns if str(c).strip() and not str(c).startswith("Unnamed")]
            if ext_cols:
                st.session_state["brand_extracted_columns"] = ext_cols

            # 새로운 파일이 업로드된 순간: 세션 상태 갱신 후 즉시 반영
            if st.session_state.get(f"last_uploaded_brand_fname_{selected_brand_key}") != brand_excel_file.name:
                st.session_state[f"last_uploaded_brand_fname_{selected_brand_key}"] = brand_excel_file.name
                clean_brand = re.sub(r'[\d=\-_월일]+', '', brand_excel_file.name).replace("발주서", "").replace("양식", "").replace(".xlsx", "").replace(".xls", "").strip()
                if clean_brand and selected_brand_key == "+ 새 브랜드사 템플릿 등록":
                    st.session_state["input_new_brand_tmpl"] = clean_brand
                if ext_cols:
                    st.session_state[txt_area_key] = ", ".join(ext_cols)
                st.rerun()
        except Exception as e:
            st.error(f"브랜드사 엑셀 파일 읽기 오류: {e}")

    # 현재 추출된 컬럼 확인 및 미리보기 뱃지 출력
    extracted_cols = st.session_state.get("brand_extracted_columns", [])
    if brand_excel_file and extracted_cols:
        st.success(f"✨ 엑셀 파일에서 총 **{len(extracted_cols)}개**의 컬럼을 자동으로 추출했습니다!")
        pills_a = "".join([f"<span style='background-color:#E2E8F0; color:#1E3D59; padding:4px 9px; border-radius:4px; margin-right:6px; margin-bottom:6px; display:inline-block; font-size:0.85rem; font-weight:bold;'>{idx+1}. {c}</span>" for idx, c in enumerate(extracted_cols)])
        st.markdown(pills_a, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### 📋 방법 B: 엑셀에서 복사하여 직접 붙여넣기 (쉼표 안 찍어도 자동 분리!)")
    st.caption("엑셀에서 헤더 행을 복사(Ctrl+C)해서 그대로 붙여넣으세요. 탭(Tab), 엔터, 쉼표 구분자를 모두 자동으로 똑똑하게 분리합니다.")

    # 텍스트 에어리어 기본값: 템플릿 컬럼이 있으면 채우고, 새 템플릿이거나 비어있으면 공란("")
    if txt_area_key not in st.session_state:
        st.session_state[txt_area_key] = ", ".join(init_cols) if init_cols else ""

    raw_header_input = st.text_area(
        "컬럼 텍스트 붙여넣기 (쉼표, 탭, 줄바꿈 모두 자동 인식)",
        height=90,
        key=txt_area_key,
        placeholder="엑셀에서 컬럼 헤더 행을 복사(Ctrl+C)하여 여기에 붙여넣으세요. (예: 쇼핑몰명, 주문번호, 수취인, 주소, 상품명, 수량)"
    )

    # 최종 컬럼 목록 결정
    text_cols = parse_header_text(raw_header_input)
    if text_cols:
        parsed_columns = text_cols
    elif extracted_cols:
        parsed_columns = extracted_cols
    else:
        parsed_columns = init_cols

    st.markdown("##### 🔍 인식된 최종 컬럼 순서 미리보기:")
    if parsed_columns:
        pills_html = "".join([f"<span style='background-color:#E2E8F0; color:#1E3D59; padding:4px 9px; border-radius:4px; margin-right:6px; margin-bottom:6px; display:inline-block; font-size:0.85rem; font-weight:bold;'>{idx+1}. {c}</span>" for idx, c in enumerate(parsed_columns)])
        st.markdown(pills_html, unsafe_allow_html=True)
    else:
        st.info("💡 아직 등록된 컬럼이 없습니다. (방법 A에서 브랜드사 발주서 엑셀 파일을 올리거나, 방법 B에 헤더 컬럼명을 붙여넣으세요)")

    st.markdown("---")
    st.markdown("#### 🚚 이 브랜드사의 배송비 & 묶음배송 정책 설정")
    st.caption("브랜드마다 상이한 묶음배송 판정 기준(체크박스)과 무료배송 기준, 기본/제주 배송비를 설정합니다. 템플릿 저장 시 함께 안전하게 보관됩니다.")

    sc_col1, sc_col2 = st.columns([1, 1])
    with sc_col1:
        st.markdown("**📦 묶음배송 판정 기준 (체크된 항목이 모두 동일할 때 1건으로 묶음)**")
        st.caption("선택한 항목들을 조합하여 동일 고객의 복수 주문을 식별하고 합포장 처리합니다.")
        k_order = f"ship_cb_order_{selected_brand_key}"
        k_name = f"ship_cb_name_{selected_brand_key}"
        k_addr = f"ship_cb_addr_{selected_brand_key}"
        k_phone = f"ship_cb_phone_{selected_brand_key}"

        if k_order not in st.session_state:
            st.session_state[k_order] = ("주문번호" in saved_targets)
        if k_name not in st.session_state:
            st.session_state[k_name] = ("수취인명" in saved_targets)
        if k_addr not in st.session_state:
            st.session_state[k_addr] = ("주소지" in saved_targets)
        if k_phone not in st.session_state:
            st.session_state[k_phone] = ("연락처" in saved_targets)

        cb_order = st.checkbox("주문번호", key=k_order)
        cb_name = st.checkbox("수취인명 (기본)", key=k_name)
        cb_addr = st.checkbox("주소지 (기본)", key=k_addr)
        cb_phone = st.checkbox("연락처 (기본)", key=k_phone)

    with sc_col2:
        st.markdown("**💰 배송비 및 무료배송 기준**")
        k_free = f"ship_free_{selected_brand_key}"
        k_base = f"ship_base_{selected_brand_key}"
        k_jeju = f"ship_jeju_{selected_brand_key}"

        if k_free not in st.session_state:
            st.session_state[k_free] = int(saved_free)
        if k_base not in st.session_state:
            st.session_state[k_base] = int(saved_base)
        if k_jeju not in st.session_state:
            st.session_state[k_jeju] = int(saved_jeju)

        cur_free = st.number_input("무료배송 기준 금액 (원 이상)", min_value=0, step=5000, key=k_free)
        cur_base = st.number_input("기본 배송비 (원)", min_value=0, step=500, key=k_base)
        cur_jeju = st.number_input("제주/도서산간 추가배송비 (원)", min_value=0, step=500, key=k_jeju)
        st.caption("💡 3단계에서 매칭한 판매단가로 묶음 건당 주문금액을 합산하여 무료배송 여부를 자동 판정합니다.")

    active_b_targets = []
    if cb_order: active_b_targets.append("주문번호")
    if cb_name: active_b_targets.append("수취인명")
    if cb_addr: active_b_targets.append("주소지")
    if cb_phone: active_b_targets.append("연락처")
    st.session_state["active_shipping_config"] = {
        "bundle_targets": active_b_targets,
        "free_shipping_threshold": cur_free,
        "base_shipping_fee": cur_base,
        "jeju_extra_fee": cur_jeju
    }

    if selected_brand_key != "+ 새 브랜드사 템플릿 등록":
        cur_b_name = selected_brand_key
        st.session_state.active_brand = cur_b_name
        st.session_state.active_brand_template_key = cur_b_name
        st.session_state.active_brand_cols = parsed_columns
    else:
        cur_b_name = target_brand_name.strip()
        if cur_b_name:
            st.session_state.active_brand = cur_b_name
            st.session_state.active_brand_template_key = cur_b_name
            st.session_state.active_brand_cols = parsed_columns
        else:
            st.session_state.active_brand_cols = parsed_columns

    def save_brand_tmpl_action():
        tmpl_k = target_brand_name.strip()
        if tmpl_k and parsed_columns:
            ship_cfg = {
                "bundle_targets": active_b_targets,
                "free_shipping_threshold": cur_free,
                "base_shipping_fee": cur_base,
                "jeju_extra_fee": cur_jeju
            }
            save_brand_template(tmpl_k, parsed_columns, brand_name=tmpl_k, shipping_config=ship_cfg)
            st.session_state.active_brand = tmpl_k
            st.session_state.active_brand_template_key = tmpl_k
            st.session_state.active_brand_cols = parsed_columns
            st.session_state["active_shipping_config"] = ship_cfg
            st.session_state["pending_brand_choice"] = tmpl_k
            st.session_state["last_brand_sync_context"] = None
            st.session_state.pop("brand_extracted_columns", None)
            st.session_state.pop("input_new_brand_tmpl", None)
            st.success(f"✨ 브랜드사 템플릿 '{tmpl_k}' 및 배송 정책이 성공적으로 저장되었습니다!")
            st.rerun()
        else:
            st.error("브랜드 템플릿 이름과 1개 이상의 컬럼이 필요합니다.")

    st.write(" ")
    btn_s1, btn_s2 = st.columns([1, 1.3])
    with btn_s1:
        if st.button("💾 이 브랜드사 템플릿 저장", type="primary", key="btn_save_brand_tmpl_main", use_container_width=True):
            save_brand_tmpl_action()
    with btn_s2:
        if st.button("👉 2단계 셀러 발주서 업로드 & 헤더 매칭으로 이동", type="primary", key="btn_goto_step2_from1", use_container_width=True):
            tmpl_k = target_brand_name.strip()
            if tmpl_k and parsed_columns:
                ship_cfg = {
                    "bundle_targets": active_b_targets,
                    "free_shipping_threshold": cur_free,
                    "base_shipping_fee": cur_base,
                    "jeju_extra_fee": cur_jeju
                }
                save_brand_template(tmpl_k, parsed_columns, brand_name=tmpl_k, shipping_config=ship_cfg)
                st.session_state.active_brand = tmpl_k
                st.session_state.active_brand_template_key = tmpl_k
                st.session_state.active_brand_cols = parsed_columns
                st.session_state["active_shipping_config"] = ship_cfg
                st.session_state["pending_brand_choice"] = tmpl_k
                st.session_state["last_brand_sync_context"] = None
                st.session_state.pop("brand_extracted_columns", None)
                st.session_state.pop("input_new_brand_tmpl", None)
            st.session_state["active_step_idx"] = 1
            st.rerun()

    st.info(f"현재 작업에 적용된 브랜드 템플릿: **{st.session_state.active_brand}** (총 {len(st.session_state.active_brand_cols)}개 컬럼 출력 예정)")

    # 템플릿 파일 백업 & 빈 양식 다운로드 / 가져오기
    with st.expander("📁 브랜드 양식 템플릿 백업 & 빈 엑셀 헤더 다운로드 / 가져오기", expanded=False):
        b_down_col1, b_down_col2 = st.columns(2)
        with b_down_col1:
            brand_json_backup = json.dumps(brand_templates, ensure_ascii=False, indent=2)
            st.download_button(
                label="📥 전체 브랜드 양식 템플릿 백업 (.json)",
                data=brand_json_backup,
                file_name=f"brand_templates_backup_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json",
                use_container_width=True
            )
        with b_down_col2:
            if st.session_state.active_brand_cols:
                empty_excel_bytes = create_empty_brand_template_excel(st.session_state.active_brand_cols, st.session_state.active_brand)
                st.download_button(
                    label=f"📑 [{st.session_state.active_brand}] 빈 엑셀 헤더 양식 다운로드 (.xlsx)",
                    data=empty_excel_bytes,
                    file_name=f"{st.session_state.active_brand}_빈양식_헤더_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        uploaded_b_json = st.file_uploader("📤 브랜드 템플릿 JSON 파일 가져오기 (추가/병합)", type=["json", "JSON"], key="uploader_brand_tmpl_json")
        if uploaded_b_json is not None:
            if st.button("🔄 브랜드 템플릿 병합 및 저장", type="primary", key="btn_save_merge_b_json"):
                try:
                    imported_b = json.loads(uploaded_b_json.read().decode("utf-8"))
                    curr_b = load_brand_templates()
                    curr_b.update(imported_b)
                    save_brand_templates(curr_b)
                    st.success("🎉 브랜드 템플릿이 성공적으로 추가/업데이트되었습니다!")
                    st.rerun()
                except Exception as ex:
                    st.error(f"오류: {ex}")



# =========================================================================
# STEP 2: 셀러 발주서 업로드 & 헤더 매칭
# =========================================================================
elif cur_step_idx == 1:
    st.markdown("""
    <div class="notion-step-banner">2️⃣ 2단계. 셀러 발주서 업로드 & 헤더 매칭</div>
    <div class="notion-step-desc-big">인플루언서/셀러로부터 받은 발주서를 등록하고, 표준 항목과 1:1로 매칭하세요. 셀러 템플릿으로 등록해두면 다음 공구 때 헤더 매칭과 품목 단가가 1초 만에 자동 로드됩니다.</div>
    """, unsafe_allow_html=True)

    seller_mappings = load_seller_mappings()
    saved_tmpl_keys = list(seller_mappings.keys())
    active_brand = st.session_state.get("active_brand") or (list(brand_templates.keys())[0] if brand_templates else "")

    # ---------------- 1. 상단: 셀러 템플릿 선택 & 관리 ----------------
    st.markdown("""
    <div style="font-size: 15px; font-weight: 700; color: #37352F; margin: 8px 0 4px 0;">
        📑 셀러 템플릿 선택 & 관리
    </div>
    """, unsafe_allow_html=True)
    st.caption("자주 사용하는 셀러 템플릿을 선택하거나, 아래에서 매칭 완료 후 신규 셀러 템플릿으로 등록할 수 있습니다.")

    s_box1, s_box2 = st.columns([3, 1.2])
    with s_box1:
        clean_brand = str(active_brand).strip()
        brand_tmpl_keys = [
            k for k in saved_tmpl_keys
            if str(seller_mappings[k].get("brand_name", "")).strip() == clean_brand
            or clean_brand in str(seller_mappings[k].get("brand_name", "")).strip()
            or str(seller_mappings[k].get("brand_name", "")).strip() in clean_brand
        ] if clean_brand else saved_tmpl_keys

        tmpl_choices = ["+ 새 셀러 템플릿 등록 (현재 파일 기준)"] + brand_tmpl_keys

        # 1. 펜딩된 셀러 선택값 처리
        if "pending_seller_tmpl_choice" in st.session_state:
            pending_seller = st.session_state.pop("pending_seller_tmpl_choice")
            st.session_state["saved_active_seller_tmpl_choice"] = pending_seller
            st.session_state["sb_seller_tmpl_choice"] = pending_seller

        # 2. 이전에 지정/선택했던 셀러 템플릿 기억값 불러오기 ⭐
        remembered_choice = (
            st.session_state.get("saved_active_seller_tmpl_choice")
            or st.session_state.get("active_seller_tmpl_key")
            or st.session_state.get("active_seller_name")
        )

        # 위젯이 다른 단계 다녀와서 언마운트되었을 때 기억값으로 안전 복원
        if "sb_seller_tmpl_choice" not in st.session_state and remembered_choice in tmpl_choices:
            st.session_state["sb_seller_tmpl_choice"] = remembered_choice

        curr_seller_val = st.session_state.get("sb_seller_tmpl_choice", remembered_choice or tmpl_choices[0])
        if curr_seller_val not in tmpl_choices:
            curr_seller_val = tmpl_choices[0]
        def_seller_idx = tmpl_choices.index(curr_seller_val)

        selected_seller_tmpl = st.selectbox(
            "적용할 셀러 템플릿 선택",
            tmpl_choices,
            index=def_seller_idx,
            key="sb_seller_tmpl_choice"
        )
        # 선택한 템플릿을 세션에 항상 영구 유지 ⭐
        st.session_state["saved_active_seller_tmpl_choice"] = selected_seller_tmpl
        if selected_seller_tmpl != "+ 새 셀러 템플릿 등록 (현재 파일 기준)":
            st.session_state["active_seller_tmpl_key"] = selected_seller_tmpl
            st.session_state["active_seller_name"] = seller_mappings.get(selected_seller_tmpl, {}).get("seller_name", selected_seller_tmpl)
    with s_box2:
        if selected_seller_tmpl != "+ 새 셀러 템플릿 등록 (현재 파일 기준)":
            st.write(" ")
            st.write(" ")
            if st.button("🗑️ 이 셀러 템플릿 삭제", key="btn_del_seller_tmpl", type="secondary"):
                delete_seller_mapping(selected_seller_tmpl)
                st.warning(f"'{selected_seller_tmpl}' 템플릿이 삭제되었습니다.")
                st.rerun()

    if selected_seller_tmpl != "+ 새 셀러 템플릿 등록 (현재 파일 기준)" and selected_seller_tmpl in seller_mappings:
        info_t = seller_mappings[selected_seller_tmpl]
        st.info(f"📌 선택된 템플릿: **{selected_seller_tmpl}** (브랜드: `{info_t.get('brand_name', '-')}` | 최종수정: `{info_t.get('updated_at', '-')}`)")

    # 셀러 템플릿 백업 & 가져오기 expander
    with st.expander("📁 셀러 매핑 템플릿 백업 및 가져오기", expanded=False):
        s_down_col1, s_down_col2 = st.columns(2)
        with s_down_col1:
            seller_json_backup = json.dumps(seller_mappings, ensure_ascii=False, indent=2)
            st.download_button(
                label="📥 전체 셀러 매핑 템플릿 백업 (.json)",
                data=seller_json_backup,
                file_name=f"seller_mappings_backup_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json",
                use_container_width=True
            )
        with s_down_col2:
            uploaded_s_json = st.file_uploader("📤 셀러 매핑 JSON 파일 복원/병합", type=["json", "JSON"], key="uploader_seller_tmpl_json")
            if uploaded_s_json is not None:
                st.markdown('<div class="save-template-btn">', unsafe_allow_html=True)
                if st.button("🔄 셀러 매핑 병합 및 저장", type="primary", key="btn_merge_s_json"):
                    try:
                        imported_s = json.loads(uploaded_s_json.read().decode("utf-8"))
                        curr_s = load_seller_mappings()
                        curr_s.update(imported_s)
                        save_seller_mappings(curr_s)
                        st.success("🎉 셀러 매핑 템플릿이 성공적으로 복원되었습니다!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"오류: {ex}")
                st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    s_col_mode, s_col_tip = st.columns([2.8, 3.2], vertical_alignment="center")
    with s_col_mode:
        upload_mode = st.radio(
            "📂 파일 업로드 방식 선택",
            ["📁 일반 업로드 (모바일 / 스마트폰 추천 ⭐)", "📑 다중 파일 일괄 병합 (PC 환경 추천 💻)"],
            horizontal=True,
            index=0,
            key="seller_upload_mode",
            help="스마트폰이나 태블릿 등 모바일 브라우저(Safari, Chrome)에서는 '일반 업로드' 모드가 가장 안정적입니다."
        )
    with s_col_tip:
        st.caption("💡 **모바일 안내**: 스마트폰(아이폰/갤럭시)에서는 카카오톡/메일의 첨부파일을 기기에 다운로드 후 '일반 업로드'로 선택하시면 오류 없이 즉시 업로드됩니다.")

    if upload_mode == "📁 일반 업로드 (모바일 / 스마트폰 추천 ⭐)":
        single_file = st.file_uploader(
            "작업할 셀러 발주서 엑셀 파일을 업로드하세요 (.xlsx, .xls, .csv)",
            type=["xlsx", "xls", "csv", "XLSX", "XLS", "CSV"],
            accept_multiple_files=False,
            key="seller_uploader_single",
            help="스마트폰 및 단일 파일 작업에 최적화된 업로더입니다."
        )
        uploaded_files = [single_file] if single_file is not None else []
    else:
        multi_files = st.file_uploader(
            "작업할 셀러 발주서 엑셀 파일들을 업로드하세요 (복수 파일 동시 드래그 or 개별 추가 가능 ⭐)",
            type=["xlsx", "xls", "csv", "XLSX", "XLS", "CSV"],
            accept_multiple_files=True,
            key="seller_uploader_multi",
            help="여러 개의 발주서 파일을 한 번에 드래그하거나, 하나씩 추가하여 자동으로 하나의 통합 발주서로 병합할 수 있습니다."
        )
        uploaded_files = multi_files if multi_files is not None else []

    if uploaded_files:
        try:
            dfs = []
            file_summaries = []
            file_names_list = []
            for uf in uploaded_files:
                if uf is None:
                    continue
                f_name = os.path.splitext(uf.name)[0]
                if hasattr(uf, "size") and uf.size == 0:
                    st.warning(f"⚠️ `{uf.name}` 파일 크기가 0바이트입니다. 기기에 완전히 다운로드된 상태인지 확인해주세요.")
                    continue
                raw_excel_df = read_uploaded_tabular_file(uf)
                part_df = sanitize_dataframe_columns(raw_excel_df)
                if not part_df.empty:
                    dfs.append(part_df)
                    file_summaries.append(f"`{uf.name}` ({len(part_df):,}건)")
                    file_names_list.append(f_name)
                else:
                    st.warning(f"⚠️ `{uf.name}` 파일에서 데이터를 읽지 못했습니다. 빈 시트이거나 서식이 일치하지 않는지 확인해주세요.")

            if dfs:
                df_seller = pd.concat(dfs, ignore_index=True)
                df_seller = sanitize_dataframe_columns(df_seller)
                if len(file_names_list) == 1:
                    seller_file_name = file_names_list[0]
                else:
                    seller_file_name = f"{file_names_list[0]} 외 {len(file_names_list)-1}개 파일"

                st.session_state.current_seller_file = seller_file_name
                st.session_state.current_seller_df = df_seller
                st.session_state.current_seller_file_summaries = file_summaries
                st.session_state.current_seller_file_count = len(dfs)
            elif uploaded_files:
                st.error("⚠️ 선택된 파일에서 데이터를 불러오지 못했습니다. 엑셀 파일이 정상적인지 확인해주세요.")
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
            df_seller = None
    # 세션 상태 보존: 사용자가 3단계, 4단계 등으로 이동했다 돌아와도 데이터를 절대 삭제하지 않음 ⭐
    if "current_seller_df" in st.session_state and st.session_state.current_seller_df is not None:
        df_seller = sanitize_dataframe_columns(st.session_state.current_seller_df)
        st.session_state.current_seller_df = df_seller
        seller_file_name = st.session_state.current_seller_file
        excel_cols = list(df_seller.columns)

        # 파일 정보 및 세션 유지 배너 + 수동 초기화 버튼
        file_count = st.session_state.get("current_seller_file_count", 1)
        file_summaries = st.session_state.get("current_seller_file_summaries", [])
        
        info_c1, info_c2 = st.columns([3.5, 1.2], vertical_alignment="center")
        with info_c1:
            if file_count > 1 and file_summaries:
                st.markdown(f"""
                <div style="font-size: 15px; font-weight: 700; color: #37352F; margin: 10px 0 4px 0;">
                    📂 현재 작업 중인 발주서: <code style="color: #0F7B6C; background: #EEF8F6; padding: 2px 7px; border-radius: 4px; font-size: 14px;">{seller_file_name}</code> 
                    <span style="background: #E0F2FE; color: #0369A1; padding: 2px 8px; border-radius: 4px; font-size: 13px; font-weight: 700; margin-left: 6px;">총 {file_count}개 파일 통합 병합</span>
                    <span style="background: #FEF3C7; color: #92400E; padding: 2px 8px; border-radius: 4px; font-size: 13px; font-weight: 700; margin-left: 4px;">총 {len(df_seller):,}건 (세션 안전 보존 중 🔒)</span>
                </div>
                """, unsafe_allow_html=True)
                st.caption(f"🔗 **병합된 세부 파일:** {' + '.join(file_summaries)} ➔ 다른 단계를 자유롭게 오가도 내용이 유지됩니다.")
            else:
                st.markdown(f"""
                <div style="font-size: 15px; font-weight: 700; color: #37352F; margin: 10px 0 4px 0;">
                    📂 현재 작업 중인 발주서: <code style="color: #0F7B6C; background: #EEF8F6; padding: 2px 7px; border-radius: 4px; font-size: 14px;">{seller_file_name}</code>
                    <span style="background: #EEF8F6; color: #0F7B6C; padding: 2px 8px; border-radius: 4px; font-size: 13px; font-weight: 700; margin-left: 6px;">총 {len(df_seller):,}건 (세션 안전 보존 중 🔒)</span>
                </div>
                """, unsafe_allow_html=True)
                st.caption("💡 3단계, 4단계 등으로 이동하더라도 발주서 데이터와 매칭 설정이 창이 열려 있는 동안 그대로 안전하게 유지됩니다.")
        with info_c2:
            if st.button("🗑️ 발주서 데이터 초기화", key="btn_clear_uploaded_seller_df", type="secondary", use_container_width=True, help="현재 불러온 발주서 데이터를 지우고 새 파일로 처음부터 다시 작업합니다."):
                st.session_state.current_seller_df = None
                st.session_state.current_seller_file = None
                st.session_state.current_seller_file_summaries = []
                st.session_state.current_seller_file_count = 0
                st.session_state.last_seller_sync_context = None
                st.session_state.pop("user_col_map", None)
                st.rerun()
        
        with st.expander("👀 업로드된 엑셀 상위 3행 미리보기 (실제 데이터 내용 확인)", expanded=False):
            df_seller_head = df_seller.head(3).copy()
            df_seller_head.index = range(1, len(df_seller_head) + 1)
            st.dataframe(df_seller_head, use_container_width=True)

        # 템플릿 데이터 불러오기 결정
        chosen_map_data = None
        if selected_seller_tmpl != "+ 새 셀러 템플릿 등록 (현재 파일 기준)" and selected_seller_tmpl in seller_mappings:
            chosen_map_data = seller_mappings[selected_seller_tmpl]
            saved_col_map = chosen_map_data.get("column_mapping", {})
            target_tmpl_name = selected_seller_tmpl
            def_seller_name = chosen_map_data.get("seller_name", selected_seller_tmpl)
            st.success(f"✅ **'{selected_seller_tmpl}'** 템플릿의 헤더 매핑 및 품목 단가 사전({len(chosen_map_data.get('product_mappings', {}))}건)이 로드되었습니다!")
        else:
            # 신규 셀러 템플릿 등록 모드: 브랜드명(리오마레 등) 자동 주입 금지! 공란("")으로 시작 ⭐
            saved_col_map = detect_column_mapping(excel_cols)
            target_tmpl_name = ""
            def_seller_name = ""
            st.caption("💡 신규 발주서입니다. 아래 컬럼 매칭을 확인하고 필요시 셀러 템플릿 이름을 입력하여 등록하세요.")

        # 스마트 컬럼 해결: 템플릿 매핑 우선, 대소문자/공백 차이 보정, 누락 항목은 파일 기반 자동 탐색으로 스마트 보완
        auto_detected_cols = detect_column_mapping(excel_cols)
        available_options = ["(선택안함)"] + excel_cols
        resolved_col_map = {}
        for f_key in STANDARD_FIELDS:
            cand = saved_col_map.get(f_key, "(선택안함)")
            if cand in excel_cols:
                resolved_col_map[f_key] = cand
            elif cand and cand != "(선택안함)":
                cand_clean = str(cand).strip().replace(" ", "").lower()
                matched_col = None
                for c in excel_cols:
                    if str(c).strip().replace(" ", "").lower() == cand_clean:
                        matched_col = str(c)
                        break
                if matched_col:
                    resolved_col_map[f_key] = matched_col
                else:
                    det = auto_detected_cols.get(f_key, "(선택안함)")
                    resolved_col_map[f_key] = det if det in excel_cols else "(선택안함)"
            else:
                det = auto_detected_cols.get(f_key, "(선택안함)")
                resolved_col_map[f_key] = det if det in excel_cols else "(선택안함)"

        # 세션 동기화: 템플릿 선택이나 파일 변경 시 모든 selectbox 위젯 상태 동기화
        sync_context = f"{selected_seller_tmpl}___{seller_file_name}___{active_brand}"
        if st.session_state.get("last_seller_sync_context") != sync_context:
            for f_key, col_name in resolved_col_map.items():
                st.session_state[f"sel_col_{f_key}"] = col_name

            # 저장된 템플릿을 선택한 경우에만 템플릿 명칭을 넣고,
            # 신규 등록 모드에서는 사용자가 직접 입력한 값이 있으면 유지하고, 없거나 브랜드명인 경우 공란("") 유지 ⭐
            if selected_seller_tmpl != "+ 새 셀러 템플릿 등록 (현재 파일 기준)":
                st.session_state["input_tmpl_title_widget"] = target_tmpl_name
            else:
                cur_user_typed = str(st.session_state.get("input_tmpl_title_widget", "")).strip()
                if not cur_user_typed or cur_user_typed == str(active_brand).strip():
                    st.session_state["input_tmpl_title_widget"] = ""

            if chosen_map_data and chosen_map_data.get("product_mappings"):
                st.session_state.active_product_mapping = chosen_map_data.get("product_mappings", {}).copy()
            if chosen_map_data and chosen_map_data.get("product_order"):
                st.session_state["active_product_order"] = list(chosen_map_data.get("product_order", []))
            else:
                st.session_state["active_product_order"] = None
            
            # 발주 일자 및 판매처명 초기값 설정
            init_date = extract_date_from_filename(seller_file_name)
            if chosen_map_data and chosen_map_data.get("custom_order_date"):
                try:
                    init_date = datetime.strptime(chosen_map_data["custom_order_date"], "%Y-%m-%d").date()
                except Exception:
                    pass
            st.session_state["custom_order_date_step2"] = init_date
            st.session_state["force_custom_date_step2"] = chosen_map_data.get("force_custom_date", True) if chosen_map_data else True
            
            # 판매처명도 템플릿에 저장된 것이 없다면 공란("") 유지 (브랜드명 자동 주입 차단)
            init_seller = chosen_map_data.get("custom_seller_name", "") if chosen_map_data else ""
            st.session_state["custom_seller_name_step2"] = init_seller

            st.session_state["last_seller_sync_context"] = sync_context
            st.rerun()

        # 위계 하향: 15px 서브헤더
        st.markdown("""
        <div style="font-size: 15px; font-weight: 700; color: #37352F; margin: 16px 0 4px 0;">
            🏷️ 셀러 템플릿 명칭 설정
        </div>
        """, unsafe_allow_html=True)
        st.caption("등록할 템플릿의 이름을 입력하세요.")

        nm_c1, nm_c2 = st.columns([3, 1.2])
        with nm_c1:
            # 브랜드명('리오마레')이 기본값으로 들어가지 않도록 보호 ⭐
            curr_title_val = st.session_state.get("input_tmpl_title_widget", "")
            if curr_title_val == str(active_brand).strip() and selected_seller_tmpl == "+ 새 셀러 템플릿 등록 (현재 파일 기준)":
                curr_title_val = ""
                st.session_state["input_tmpl_title_widget"] = ""

            input_tmpl_title = st.text_input(
                "셀러 템플릿 이름 (저장할 이름)",
                value=curr_title_val,
                key="input_tmpl_title_widget",
                placeholder="예: 셀러명 또는 셀러_브랜드"
            )
        with nm_c2:
            st.text_input("적용 브랜드사", value=active_brand, disabled=True, key="input_brand_display_step2")

        st.session_state.active_seller_tmpl_key = input_tmpl_title.strip()
        st.session_state.active_seller_name = input_tmpl_title.strip()

        # 위계 하향: 15px 서브헤더
        st.markdown("""
        <div style="font-size: 15px; font-weight: 700; color: #37352F; margin: 18px 0 4px 0;">
            🎯 셀러 엑셀 컬럼 ➔ 표준 항목 1:1 매칭
        </div>
        """, unsafe_allow_html=True)
        st.caption("주의: 상품코드(숫자) 열이 아니라 실제 '한글 상품명' 열을 정확히 선택해주세요!")

        user_col_map = {}
        grid_cols = st.columns(3)
        fields_list = list(STANDARD_FIELDS.items())

        for idx, (f_key, f_label) in enumerate(fields_list):
            with grid_cols[idx % 3]:
                # 이전에 사용자가 매칭해둔 값이 세션에 있으면 최우선 복원 ⭐
                user_saved_map = st.session_state.get("user_col_map", {})
                if f_key in user_saved_map and user_saved_map[f_key] in available_options:
                    curr_val = user_saved_map[f_key]
                else:
                    curr_val = st.session_state.get(f"sel_col_{f_key}", resolved_col_map.get(f_key, "(선택안함)"))
                if curr_val not in available_options:
                    curr_val = "(선택안함)"
                def_idx = available_options.index(curr_val)
                
                selected_col = st.selectbox(
                    f"**{f_label}**",
                    options=available_options,
                    index=def_idx,
                    key=f"sel_col_{f_key}"
                )
                user_col_map[f_key] = selected_col

        st.session_state.user_col_map = user_col_map
        matched_count = sum(1 for v in user_col_map.values() if v and v != "(선택안함)")
        st.success(f"✨ **{len(STANDARD_FIELDS)}개 표준 항목 중 {matched_count}개가 매칭되었습니다!**")

        # ---------------- 발주 일자 & 판매처 커스텀 설정 ----------------
        st.markdown("""
        <div style="font-size: 15px; font-weight: 700; color: #37352F; margin: 18px 0 4px 0;">
            📅 발주 일자 & 🏪 판매처(쇼핑몰) 커스텀 설정
        </div>
        """, unsafe_allow_html=True)
        st.caption("일자 커스텀 및 판매처명을 직접 지정할 수 있습니다. 입력된 판매처명은 2행부터 주문 데이터가 존재하는 모든 행에 자동으로 일괄 기입됩니다.")

        cust_c1, cust_c2 = st.columns(2)
        with cust_c1:
            default_c_date = st.session_state.get("active_custom_order_date") or st.session_state.get("custom_order_date_step2") or extract_date_from_filename(seller_file_name)
            custom_order_date = st.date_input(
                "📅 **발주 기준 일자**",
                value=default_c_date,
                key="custom_order_date_step2",
                help="발주서에 표기될 일자입니다. 파일명에서 자동 추출된 일자가 기본 입력됩니다."
            )

        with cust_c2:
            default_c_seller = st.session_state.get("active_custom_seller_name")
            if default_c_seller is None or default_c_seller == "":
                default_c_seller = st.session_state.get("custom_seller_name_step2", "")
            if not default_c_seller:
                default_c_seller = extract_clean_seller_from_filename(seller_file_name)
            custom_seller_val = st.text_input(
                "🏪 **판매처 / 쇼핑몰명 직접 지정**",
                value=default_c_seller,
                key="custom_seller_name_step2",
                placeholder="예: 리오마레, 스마트스토어, 쿠팡 등",
                help="입력 시 엑셀 2행부터 주문 데이터가 존재하는 모든 행의 [판매처/쇼핑몰명]에 해당 값이 자동 반영됩니다."
            )

        sub_c1, sub_c2 = st.columns(2)
        with sub_c1:
            force_custom_date = st.checkbox(
                "📌 **전체 발주 일자를 이 날짜로 일괄 통일** (엑셀 내 시간 무관 일괄 적용)",
                value=st.session_state.get("active_force_custom_date", st.session_state.get("force_custom_date_step2", True)),
                key="force_custom_date_step2"
            )
        with sub_c2:
            st.caption("💡 특정 판매처명을 입력하시면 2행부터 주문이 있는 모든 행에 자동 추가됩니다. 비워두시면 위 매칭 컬럼(판매처)의 값이 사용됩니다.")

        st.session_state["active_custom_order_date"] = custom_order_date
        st.session_state["active_force_custom_date"] = force_custom_date
        st.session_state["active_custom_seller_name"] = custom_seller_val

        # ---------------- 실시간 브랜드사 최종 발주서 매칭 미리보기 (상위 5행) ----------------
        st.markdown(f"""
        <div style="font-size: 15px; font-weight: 700; color: #37352F; margin: 20px 0 4px 0;">
            📋 매칭 결과 미리보기 (브랜드사 [{active_brand}] 최종 발주서 양식 기준 상위 5행)
        </div>
        """, unsafe_allow_html=True)
        st.caption(f"💡 현재 1단계에서 지정된 브랜드사 **[{active_brand}]** 발주서의 컬럼 순서 및 양식 그대로 매칭된 상위 5행입니다. 어떤 위치에 어떤 정보가 들어가는지 확인하세요. (상품명/단가는 3단계에서 브랜드 공식명으로 변환됩니다)")

        brand_templates = load_brand_templates()
        active_tmpl_k = st.session_state.get("active_brand_template_key", active_brand)
        active_brand_template = brand_templates.get(active_tmpl_k) or brand_templates.get(active_brand, {})
        if "columns" not in active_brand_template or not active_brand_template.get("columns"):
            active_brand_template = {"columns": st.session_state.get("active_brand_cols", [])}
        sample_df = df_seller.head(5)
        norm_sample_5 = process_seller_dataframe(
            sample_df, 
            seller_file_name, 
            user_col_map, 
            product_mapping_dict={},
            custom_order_date=custom_order_date,
            force_custom_date=force_custom_date,
            custom_seller_name=custom_seller_val
        )
        preview_brand_df = build_brand_order_sheet(norm_sample_5, active_brand_template)
        preview_clean_df = preview_brand_df.drop(columns=[c for c in ["_row_highlight", "_row_color"] if c in preview_brand_df.columns])
        preview_clean_df.index = range(1, len(preview_clean_df) + 1)
        st.dataframe(preview_clean_df, use_container_width=True, height=210)

        st.markdown("---")
        m_col1, m_col2, m_col3 = st.columns([1.5, 2, 1.2])
        with m_col1:
            if st.button("💾 이 셀러 템플릿 등록하기", type="primary", key="btn_save_seller_tmpl_main", use_container_width=True):
                final_key = input_tmpl_title.strip()
                if not final_key:
                    st.error("셀러 템플릿 이름을 입력해주세요.")
                else:
                    curr_pm = st.session_state.get("active_product_mapping", {})
                    if not curr_pm and chosen_map_data:
                        curr_pm = chosen_map_data.get("product_mappings", {})
                    
                    c_date_str = custom_order_date.strftime("%Y-%m-%d") if hasattr(custom_order_date, "strftime") else str(custom_order_date)
                    save_seller_mapping(
                        final_key, final_key, active_brand, user_col_map, curr_pm,
                        custom_seller_name=custom_seller_val,
                        custom_order_date=c_date_str,
                        force_custom_date=force_custom_date
                    )
                    st.session_state["last_seller_sync_context"] = f"{final_key}___{seller_file_name}___{active_brand}"
                    st.session_state["pending_seller_tmpl_choice"] = final_key
                    st.session_state["saved_active_seller_tmpl_choice"] = final_key
                    st.session_state["active_seller_tmpl_key"] = final_key
                    st.session_state["active_seller_name"] = final_key
                    st.success(f"✨ '{final_key}' 셀러 템플릿이 성공적으로 등록되었습니다!")
                    st.rerun()

        with m_col2:
            if st.button("👉 3단계 품목 매칭 & 본사 판매단가 설정으로 이동", type="primary", key="btn_goto_step3_from2", use_container_width=True):
                # 3단계로 이동하기 전 현재 지정된 셀러 템플릿과 명칭을 완벽 보존 ⭐
                st.session_state["saved_active_seller_tmpl_choice"] = selected_seller_tmpl
                st.session_state["active_seller_tmpl_key"] = input_tmpl_title.strip() or selected_seller_tmpl
                st.session_state["active_seller_name"] = input_tmpl_title.strip() or selected_seller_tmpl
                st.session_state["active_step_idx"] = 2
                st.rerun()

        with m_col3:
            if st.button("🔄 매핑 재감지/초기화", type="secondary", key="btn_reset_seller_col_map", use_container_width=True):
                st.session_state["last_seller_sync_context"] = None
                st.rerun()
    else:
        st.info("👆 셀러 발주서 엑셀 파일을 업로드해주세요.")
        nav_e1, nav_e2 = st.columns([1, 1.3])
        with nav_e1:
            if st.button("👈 1단계 브랜드 양식으로 돌아가기", type="secondary", use_container_width=True, key="btn_goto_step1_from2"):
                st.session_state["active_step_idx"] = 0
                st.rerun()
        with nav_e2:
            if "current_seller_df" in st.session_state and st.session_state.current_seller_df is not None:
                if st.button("👉 3단계 품목 매칭 & 본사 판매단가 설정으로 이동", type="primary", use_container_width=True, key="btn_goto_step3_from2_prev"):
                    st.session_state["active_step_idx"] = 2
                    st.rerun()

# =========================================================================
# STEP 3: 상품명 & 옵션 & 판매단가 매칭
# =========================================================================
elif cur_step_idx == 2:
    st.markdown("""
    <div class="notion-step-banner">3️⃣ 3단계. 품목 매칭 & 본사 판매단가 설정</div>
    <div class="notion-step-desc-big">셀러의 상품명/옵션을 브랜드사 공식 품목으로 변환하고 본사 판매단가를 지정하세요. 단가가 기입되어야 일자별 수량표에 매출이 자동 집계됩니다.</div>
    """, unsafe_allow_html=True)

    if "current_seller_df" in st.session_state and "user_col_map" in st.session_state:
        df_seller = st.session_state.current_seller_df
        seller_file_name = st.session_state.current_seller_file
        user_col_map = st.session_state.user_col_map

        unique_items = extract_unique_products(df_seller, user_col_map)

        if not unique_items:
            st.warning("⚠️ 2단계에서 [상품명/품목] 컬럼이 선택되지 않았거나 데이터가 비어 있습니다. 2단계로 돌아가 상품명 열을 지정해주세요.")
        else:
            st.markdown(f"**현재 파일에서 발견된 고유 상품/옵션 종류:** 총 `{len(unique_items)}`건")
            st.info("💡 **무료배송 판정 원리:** 여기서 입력하신 **'본사 판매단가' × 수량**으로 각 주문건별 총 구매금액을 자동 계산하여, 4번 사이드바의 무료배송 기준(예: 5만원) 이상인지 자동으로 판정합니다.")

            active_tmpl_key = st.session_state.get("active_seller_tmpl_key", seller_file_name)
            active_seller_tmpl = st.session_state.get("active_seller_tmpl_key", active_tmpl_key)
            seller_map_data = get_seller_mapping(active_tmpl_key) or get_seller_mapping(seller_file_name) or {}
            active_brand = st.session_state.get("active_brand") or seller_map_data.get("brand_name") or "리오마레"
            active_seller_name = st.session_state.get("active_seller_name", active_tmpl_key or "")

            if "active_product_mapping" in st.session_state and st.session_state.active_product_mapping:
                saved_pm = st.session_state.active_product_mapping
            else:
                saved_pm = seller_map_data.get("product_mappings", {}).copy()
                st.session_state.active_product_mapping = saved_pm

            if "active_product_order" not in st.session_state or st.session_state.active_product_order is None:
                st.session_state.active_product_order = list(seller_map_data.get("product_order", []))

            # 브랜드 공식 매핑 풀 수집
            known_brand_pool = get_all_known_brand_mappings(active_brand)

            # 세션에 이전 raw_p 기본값으로 오염된 매핑이 있거나, 미매칭된 품목이 있다면 정리 및 공식 브랜드 풀 자동 매칭
            unmatched_or_polluted = [
                (u_p, u_o) for u_p, u_o in unique_items
                if f"{u_p}||{u_o}" not in saved_pm
                or not saved_pm.get(f"{u_p}||{u_o}", {}).get("brand_product")
                or saved_pm.get(f"{u_p}||{u_o}", {}).get("brand_product") == u_p
                or "에이핏" in str(saved_pm.get(f"{u_p}||{u_o}", {}).get("brand_product", ""))
            ]
            if unmatched_or_polluted and known_brand_pool:
                pre_matched, pre_cnt = auto_match_products_against_brand_pool(unmatched_or_polluted, known_brand_pool)
                if pre_matched:
                    saved_pm.update(pre_matched)
                    st.session_state.active_product_mapping = saved_pm

            # 혹시 여전히 brand_product에 셀러 원본명(raw_p)이나 에이핏이 남아있는 항목이 있다면 공백으로 초기화 (단, 변경없이 진행 모드일 때는 보존)
            if not st.session_state.get("step3_keep_original_mode"):
                for u_p, u_o in unique_items:
                    u_pair = f"{u_p}||{u_o}"
                    cur_inf = saved_pm.get(u_pair)
                    if cur_inf and not cur_inf.get("keep_original"):
                        cur_bp = str(cur_inf.get("brand_product", "")).strip()
                        if cur_bp == u_p or "에이핏" in cur_bp:
                            cur_inf["brand_product"] = ""
                            cur_inf["is_matched"] = False

            # ----------------- 1. 기존 발주서 품목명 변경없이 진행 (원클릭 그대로 적용) -----------------
            is_keep_orig_active = bool(st.session_state.get("step3_keep_original_mode", False))

            with st.expander("📦 기존 발주서 품목명 변경없이 진행 (셀러 상품명 그대로 적용)", expanded=True):
                st.markdown("""
                별도의 브랜드 공식 상품명으로 변경하지 않고, **기존 발주서의 셀러 상품명을 그대로 사용하여 진행**합니다.<br>
                버튼을 누르면 셀러 상품명이 공식 상품명에 100% 동일하게 일괄 적용되며, 공식 품목 마스터에는 등록되지 않습니다.
                """, unsafe_allow_html=True)

                if is_keep_orig_active:
                    st.info("🔵 **현재 '기존 발주서 품목명 변경없이 진행 (셀러 상품명 유지)' 모드가 적용되어 있습니다.** (좌측 단계에 '🔵 변경없음' 표시)")
                    btn_c1, btn_c2 = st.columns([3, 1.2])
                    with btn_c1:
                        if st.button("👉 기존 발주서 품목명 변경없이 진행 (재적용)", type="primary", use_container_width=True, key="btn_keep_orig_reapply"):
                            st.session_state["step3_keep_original_mode"] = True
                            for s_idx, (u_p, u_o) in enumerate(unique_items):
                                u_pair = f"{u_p}||{u_o}"
                                prev_inf = saved_pm.get(u_pair, {})
                                saved_pm[u_pair] = {
                                    "brand_product": u_p,
                                    "brand_option": u_o if u_o else "",
                                    "unit_price": prev_inf.get("unit_price", 0),
                                    "is_matched": True,
                                    "keep_original": True,
                                    "sort_order": prev_inf.get("sort_order", s_idx + 1)
                                }
                            st.session_state.active_product_mapping = saved_pm
                            st.success("🎉 기존 발주서 품목명이 그대로 적용되었습니다! (공식 품목 마스터 미등록)")
                            st.rerun()
                    with btn_c2:
                        if st.button("↩️ 모드 해제", use_container_width=True, key="btn_cancel_keep_orig"):
                            st.session_state["step3_keep_original_mode"] = False
                            for u_p, u_o in unique_items:
                                u_pair = f"{u_p}||{u_o}"
                                if saved_pm.get(u_pair, {}).get("keep_original"):
                                    saved_pm[u_pair]["brand_product"] = ""
                                    saved_pm[u_pair]["is_matched"] = False
                                    saved_pm[u_pair]["keep_original"] = False
                            st.session_state.active_product_mapping = saved_pm
                            st.rerun()
                else:
                    if st.button("👉 기존 발주서 품목명 변경없이 진행", type="primary", use_container_width=True, key="btn_keep_original_seller_items"):
                        st.session_state["step3_keep_original_mode"] = True
                        for s_idx, (u_p, u_o) in enumerate(unique_items):
                            u_pair = f"{u_p}||{u_o}"
                            prev_inf = saved_pm.get(u_pair, {})
                            saved_pm[u_pair] = {
                                "brand_product": u_p,
                                "brand_option": u_o if u_o else "",
                                "unit_price": prev_inf.get("unit_price", 0),
                                "is_matched": True,
                                "keep_original": True,
                                "sort_order": prev_inf.get("sort_order", s_idx + 1)
                            }
                        st.session_state.active_product_mapping = saved_pm
                        active_seller_tmpl = st.session_state.get("active_seller_tmpl_key")
                        all_mappings = load_seller_mappings()
                        if active_seller_tmpl and active_seller_tmpl in all_mappings:
                            save_seller_mapping(active_seller_tmpl, active_seller_name, active_brand, user_col_map, saved_pm)
                        # 공식 품목 마스터(brand_product_master)에는 업데이트하지 않음 (요구사항 반영)
                        st.success("🎉 기존 발주서 품목명이 그대로 적용되었습니다! (공식 품목 마스터에는 등록되지 않습니다)")
                        st.rerun()

            # 빠른 단가 일괄 불러오기 (엑셀 업로드 & 텍스트 복붙)
            with st.expander("⚡ 단가표 일괄 불러오기 (엑셀 업로드 or 텍스트 복붙)", expanded=False):
                st.markdown("본사 단가표 엑셀 파일(`.xlsx`)을 올리시거나 텍스트를 붙여넣으시면, **옵션코드(104469-0017, 104012 등)**와 단가를 자동으로 대조하여 전체 품목의 판매단가를 1초 만에 채웁니다.")

                tab_excel, tab_paste, tab_manual = st.tabs([
                    "📂 1. 단가표 엑셀 파일 업로드 (가장 확실 ⭐)", 
                    "📋 2. 텍스트 복사-붙여넣기",
                    "➕ 3. 공식 품목 직접 1건 사전등록"
                ])

                # ----------------- 탭 1: 단가표 엑셀 파일 업로드 -----------------
                with tab_excel:
                    price_excel_file = st.file_uploader(
                        "본사 단가표/가격표 엑셀 파일 업로드 (.xlsx, .xls, .csv)",
                        type=["xlsx", "xls", "csv", "XLSX", "XLS", "CSV"],
                        key="price_file_uploader_v2"
                    )

                    if price_excel_file:
                        try:
                            df_px = read_uploaded_tabular_file(price_excel_file)
                            df_px_head = df_px.head(3).copy()
                            df_px_head.index = range(1, len(df_px_head) + 1)
                            st.dataframe(df_px_head, use_container_width=True)

                            px_cols = list(df_px.columns)

                            # 공식 상품명 열 추정
                            def_name_col = "(코드/식별열과 동일)"
                            for c in px_cols:
                                c_str = str(c).strip()
                                if any(kw in c_str for kw in ["브랜드_공식상품명", "공식상품명", "공식명", "상품명", "품목명", "제품명", "상품"]):
                                    def_name_col = c
                                    break
                            name_col_options = ["(코드/식별열과 동일)"] + px_cols

                            # 코드 열 추정
                            def_code_col = px_cols[0]
                            for c in px_cols:
                                c_str = str(c).strip()
                                if any(kw in c_str for kw in ["코드", "옵션코드", "상품코드", "품목코드", "자체코드"]):
                                    def_code_col = c
                                    break

                            # 단가 열 추정
                            def_price_col = px_cols[-1]
                            for c in px_cols:
                                c_str = str(c).strip()
                                if any(kw in c_str for kw in ["판매가", "단가", "금액", "가격", "공급가", "소비자가"]):
                                    def_price_col = c
                                    break

                            col_sel1, col_sel2, col_sel3 = st.columns([1.3, 1.2, 1.1])
                            with col_sel1:
                                chosen_name_col = st.selectbox(
                                    "🏷️ 브랜드 공식 상품명 열 (선택)",
                                    name_col_options,
                                    index=name_col_options.index(def_name_col) if def_name_col in name_col_options else 0,
                                    key="chosen_name_col_sb",
                                    help="단가표에 공식 상품명 열이 따로 있다면 선택하세요. 없을 시 코드/식별열이 상품명으로 사용됩니다."
                                )
                            with col_sel2:
                                chosen_code_col = st.selectbox(
                                    "🎯 옵션코드 / 상품 식별 열 선택",
                                    px_cols,
                                    index=px_cols.index(def_code_col) if def_code_col in px_cols else 0,
                                    key="chosen_code_col_sb"
                                )
                            with col_sel3:
                                chosen_price_col = st.selectbox(
                                    "💰 판매단가(금액) 열 선택",
                                    px_cols,
                                    index=px_cols.index(def_price_col) if def_price_col in px_cols else len(px_cols) - 1,
                                    key="chosen_price_col_sb"
                                )

                            if st.button("⚡ 이 엑셀 단가표로 발주서 매칭 & 공식 품목 일괄 사전등록", type="primary", key="btn_apply_excel_px"):
                                matched_count = 0
                                matched_excel_indices = set()
                                excel_items_list = []
                                excel_price_dict = {}
                                excel_order_dict = {}
                                excel_name_map = {}

                                for r_idx, (_, r) in enumerate(df_px.iterrows()):
                                    raw_c = str(r.get(chosen_code_col, "")).strip()
                                    if raw_c == "nan":
                                        raw_c = ""
                                    raw_p = str(r.get(chosen_price_col, "0")).strip()
                                    clean_p_num = re.sub(r'[^0-9]', '', raw_p)
                                    p_val = int(clean_p_num) if clean_p_num else 0
                                    
                                    raw_name = ""
                                    if chosen_name_col != "(코드/식별열과 동일)":
                                        raw_name = str(r.get(chosen_name_col, "")).strip()
                                        if raw_name == "nan":
                                            raw_name = ""
                                    disp_name = raw_name if raw_name else raw_c
                                    if not disp_name and not raw_c:
                                        continue

                                    item_entry = {
                                        "index": r_idx,
                                        "code": raw_c,
                                        "name": disp_name,
                                        "price": p_val,
                                        "order": r_idx + 1
                                    }
                                    excel_items_list.append(item_entry)

                                    if raw_c:
                                        excel_price_dict[raw_c] = item_entry
                                        excel_order_dict[raw_c] = r_idx + 1
                                        excel_price_dict[raw_c.replace("-", "")] = item_entry
                                        excel_order_dict[raw_c.replace("-", "")] = r_idx + 1
                                    if disp_name:
                                        excel_name_map[disp_name] = item_entry

                                # 발주서 품목과 대조
                                for u_p, u_o in unique_items:
                                    k_pair = f"{u_p}||{u_o}"
                                    found_entry = None
                                    clean_uo = str(u_o).strip()

                                    # 1) 셀러 옵션명(u_o) 대조
                                    if clean_uo and clean_uo in excel_price_dict:
                                        found_entry = excel_price_dict[clean_uo]
                                    elif clean_uo and clean_uo.replace("-", "") in excel_price_dict:
                                        found_entry = excel_price_dict[clean_uo.replace("-", "")]
                                    
                                    # 2) 셀러 상품명(u_p) 안의 코드 대조
                                    if found_entry is None:
                                        for c_key, c_entry in excel_price_dict.items():
                                            if len(c_key) >= 4 and c_key in str(u_p):
                                                found_entry = c_entry
                                                break

                                    # 3) 공식 상품명(disp_name) 직접 대조
                                    if found_entry is None:
                                        for n_key, n_entry in excel_name_map.items():
                                            if len(n_key) >= 2 and (n_key in str(u_p) or str(u_p) in n_key):
                                                found_entry = n_entry
                                                break

                                    if found_entry is not None:
                                        matched_excel_indices.add(found_entry["index"])
                                        cand_name = found_entry["name"]
                                        curr_brand_p = cand_name if (cand_name and cand_name != u_p and "에이핏" not in cand_name) else ""
                                        if not curr_brand_p:
                                            curr_brand_p = str(saved_pm.get(k_pair, {}).get("brand_product", "")).strip()
                                            if (not curr_brand_p or curr_brand_p == u_p or "에이핏" in curr_brand_p) and k_pair in known_brand_pool:
                                                curr_brand_p = known_brand_pool[k_pair].get("brand_product", "")
                                            if curr_brand_p == u_p or "에이핏" in curr_brand_p:
                                                curr_brand_p = cand_name or ""
                                        
                                        lead_num = extract_leading_number(curr_brand_p) or extract_leading_number(u_p)
                                        s_ord = lead_num if lead_num is not None else found_entry["order"]

                                        saved_pm[k_pair] = {
                                            "brand_product": curr_brand_p,
                                            "brand_option": u_o if u_o else found_entry["code"],
                                            "code": found_entry["code"],
                                            "unit_price": found_entry["price"],
                                            "sort_order": s_ord,
                                            "is_matched": True
                                        }
                                        matched_count += 1

                                # 발주서에 없는 엑셀 단가표 품목들을 사전등록 공식품목(⏳ 매칭대기)으로 저장
                                pre_reg_count = 0
                                for e_entry in excel_items_list:
                                    if e_entry["index"] in matched_excel_indices:
                                        continue
                                    e_name = e_entry["name"]
                                    e_code = e_entry["code"]
                                    e_price = e_entry["price"]
                                    e_order = e_entry["order"]
                                    k_off = f"__official__{e_name}||{e_code}"
                                    saved_pm[k_off] = {
                                        "brand_product": e_name,
                                        "brand_option": e_code if e_code != e_name else "",
                                        "code": e_code,
                                        "unit_price": e_price,
                                        "sort_order": e_order,
                                        "is_pre_registered": True,
                                        "is_matched": False
                                    }
                                    pre_reg_count += 1

                                # 엑셀 단가표 순서 및 번호순으로 active_product_order 재정렬
                                def get_excel_sort_key(item_tuple):
                                    kp = f"{item_tuple[0]}||{item_tuple[1]}"
                                    if kp in saved_pm and "sort_order" in saved_pm[kp]:
                                        return (0, int(saved_pm[kp]["sort_order"]))
                                    return (1, 999999)

                                sorted_excel_items = sorted(unique_items, key=get_excel_sort_key)
                                new_excel_order = [f"{p}||{o}" for p, o in sorted_excel_items]
                                for k in saved_pm.keys():
                                    if k.startswith("__official__") and k not in new_excel_order:
                                        new_excel_order.append(k)
                                st.session_state.active_product_order = new_excel_order

                                # 등록된 셀러 템플릿 및 브랜드 마스터 영구 업데이트
                                active_brand = st.session_state.get("active_brand", "호무로")
                                active_seller_tmpl = st.session_state.get("active_seller_tmpl_key")
                                active_seller_name = st.session_state.get("active_seller_name", active_seller_tmpl or "")
                                all_mappings = load_seller_mappings()
                                if active_seller_tmpl and active_seller_tmpl in all_mappings:
                                    save_seller_mapping(active_seller_tmpl, active_seller_name, active_brand, user_col_map, saved_pm, product_order=new_excel_order)
                                update_brand_product_master_from_pm(active_brand, saved_pm)
                                st.session_state.active_product_mapping = saved_pm

                                for k in list(st.session_state.keys()):
                                    if k.startswith("brand_price_") or k.startswith("brand_p_") or k.startswith("brand_o_"):
                                        del st.session_state[k]

                                res_msgs = []
                                if matched_count > 0:
                                    res_msgs.append(f"발주서 품목 **{matched_count}개 자동 매칭**")
                                if pre_reg_count > 0:
                                    res_msgs.append(f"발주서 외 공식 품목 **{pre_reg_count}개 사전등록 (⏳ 매칭대기 목록)**")
                                if res_msgs:
                                    st.success(f"🎉 엑셀 단가표에서 {' 및 '.join(res_msgs)} 완료되어 공식 마스터에 영구 등록되었습니다!")
                                else:
                                    st.warning("⚠️ 엑셀 단가표에서 유효한 품목 정보를 읽지 못했습니다. 선택한 열을 확인해주세요.")
                                st.rerun()

                        except Exception as e:
                            st.error(f"엑셀 파일 처리 오류: {e}")

                # ----------------- 탭 2: 텍스트 복사-붙여넣기 -----------------
                with tab_paste:
                    bulk_text = st.text_area(
                        "단가표 텍스트 붙여넣기 (한 줄씩 복사)",
                        placeholder="인쌀라티씨메 마일드스파이시 라이스 160g 6개 34300\n올리브유 참치 80g 6개 22900\n104012 10,000",
                        height=120,
                        key="bulk_price_textarea_v2"
                    )

                    if st.button("⚡ 붙여넣은 텍스트로 공식 상품명 & 단가 일괄 등록 및 발주서 매칭", type="primary", key="btn_apply_paste_px"):
                        if bulk_text:
                            parsed_entries = []
                            for idx_line, line in enumerate(bulk_text.strip().split("\n")):
                                line_clean = line.strip()
                                if not line_clean:
                                    continue

                                # 1. 탭 또는 다중 공백 기준 분리 확인
                                if "\t" in line_clean:
                                    parts = [p.strip() for p in line_clean.split("\t") if p.strip()]
                                elif "  " in line_clean:
                                    parts = [p.strip() for p in re.split(r'\s{2,}', line_clean) if p.strip()]
                                else:
                                    parts = [line_clean]

                                price = 0
                                official_name = ""
                                official_opt = ""

                                if len(parts) >= 2:
                                    digits = re.sub(r'[^0-9]', '', parts[-1])
                                    if digits and int(digits) >= 100:
                                        price = int(digits)
                                        if len(parts) == 2:
                                            official_name = parts[0]
                                        else:
                                            official_name = parts[0]
                                            official_opt = parts[1]

                                if price == 0:
                                    # 공백으로 토큰 분리 후 마지막 숫자 추출
                                    tokens = line_clean.split()
                                    if len(tokens) >= 2:
                                        digits = re.sub(r'[^0-9]', '', tokens[-1])
                                        if digits and int(digits) >= 100:
                                            price = int(digits)
                                            official_name = " ".join(tokens[:-1]).strip()

                                if price > 0:
                                    parsed_entries.append({
                                        "index": idx_line,
                                        "official_name": official_name if official_name else line_clean,
                                        "official_opt": official_opt,
                                        "price": price,
                                        "raw": line_clean
                                    })

                            def calculate_match_score(raw_p, raw_o, entry):
                                score = 0
                                p_name = entry["official_name"]
                                p_opt = entry["official_opt"]

                                target_str = f"{raw_p} {raw_o}".strip()
                                entry_str = f"{p_name} {p_opt}".strip()

                                clean_target = re.sub(r'[^0-9a-zA-Z가-힣]', '', target_str).lower()
                                clean_entry = re.sub(r'[^0-9a-zA-Z가-힣]', '', entry_str).lower()

                                if clean_target == clean_entry:
                                    return 1000

                                # 순수 제품코드 여부 (104012, 104469-0017 등)
                                is_pure_code = bool(re.match(r'^[0-9A-Za-z\-_]+$', p_name))
                                if is_pure_code and len(p_name) >= 4 and re.search(r'\d', p_name):
                                    code_clean = re.sub(r'[^0-9a-zA-Z]', '', p_name).lower()
                                    if code_clean in re.sub(r'[^0-9a-zA-Z]', '', str(raw_o)).lower():
                                        return 900
                                    if code_clean in re.sub(r'[^0-9a-zA-Z]', '', str(raw_p)).lower():
                                        return 850

                                # 한국어 키워드 대조
                                t_ko = set(re.findall(r'[가-힣]{2,}', target_str))
                                e_ko = set(re.findall(r'[가-힣]{2,}', entry_str))

                                has_substr = (clean_target in clean_entry or clean_entry in clean_target)
                                ko_shared = t_ko.intersection(e_ko)

                                # 한글 품목명이 전혀 다른 경우 매칭 제외
                                if t_ko and e_ko and not ko_shared and not has_substr:
                                    return 0

                                if has_substr:
                                    score += 300
                                if ko_shared:
                                    score += len(ko_shared) * 50

                                # 규격/용량/수량 대조 (예: 160g vs 80g, 6개 vs 4개)
                                def get_specs(text):
                                    return set(re.findall(r'\d+(?:g|ml|kg|캔|개|입|세트|box)?', text.lower()))

                                t_specs = get_specs(target_str)
                                e_specs = get_specs(entry_str)

                                common_specs = t_specs.intersection(e_specs)
                                if common_specs:
                                    score += len(common_specs) * 40

                                for s1 in t_specs:
                                    for s2 in e_specs:
                                        if s1 != s2:
                                            u1 = re.sub(r'\d', '', s1)
                                            u2 = re.sub(r'\d', '', s2)
                                            if u1 and u1 == u2:
                                                score -= 100

                                return max(0, score)

                            matched_count = 0
                            matched_paste_indices = set()
                            for u_p, u_o in unique_items:
                                k_pair = f"{u_p}||{u_o}"
                                best_entry = None
                                best_score = 0
                                best_entry_idx = 0

                                for idx_e, e in enumerate(parsed_entries):
                                    sc = calculate_match_score(u_p, u_o, e)
                                    if sc > best_score:
                                        best_score = sc
                                        best_entry = e
                                        best_entry_idx = idx_e

                                if best_entry is not None and best_score > 0:
                                    matched_paste_indices.add(best_entry.get("index", best_entry_idx))
                                    # 사용자가 입력한 공식 품목명이 코드가 아닌 상품명일 경우 우선 반영! ⭐
                                    p_cand = best_entry["official_name"]
                                    is_pure_code = bool(re.match(r'^[0-9A-Za-z\-_]+$', p_cand))
                                    if not is_pure_code and len(p_cand.strip()) >= 2:
                                        curr_brand_p = p_cand.strip()
                                    else:
                                        curr_brand_p = str(saved_pm.get(k_pair, {}).get("brand_product", "")).strip()
                                        if (not curr_brand_p or curr_brand_p == u_p or "에이핏" in curr_brand_p) and k_pair in known_brand_pool:
                                            curr_brand_p = known_brand_pool[k_pair].get("brand_product", "")
                                        if curr_brand_p == u_p or "에이핏" in curr_brand_p:
                                            curr_brand_p = ""

                                    curr_brand_o = best_entry["official_opt"].strip() if best_entry.get("official_opt") else saved_pm.get(k_pair, {}).get("brand_option", u_o)

                                    lead_num = extract_leading_number(curr_brand_p) or extract_leading_number(u_p)
                                    s_ord = lead_num if lead_num is not None else (best_entry_idx + 1)

                                    saved_pm[k_pair] = {
                                        "brand_product": curr_brand_p,
                                        "brand_option": curr_brand_o,
                                        "unit_price": best_entry["price"],
                                        "sort_order": s_ord,
                                        "is_matched": True
                                    }
                                    matched_count += 1

                            # 발주서에 없는 텍스트 단가표 품목들을 사전등록 공식품목(⏳ 매칭대기)으로 저장
                            pre_reg_count = 0
                            for idx_e, e_entry in enumerate(parsed_entries):
                                if e_entry.get("index", idx_e) in matched_paste_indices:
                                    continue
                                e_name = e_entry["official_name"]
                                e_opt = e_entry.get("official_opt", "")
                                e_price = e_entry["price"]
                                k_off = f"__official__{e_name}||{e_opt}"
                                saved_pm[k_off] = {
                                    "brand_product": e_name,
                                    "brand_option": e_opt,
                                    "code": e_opt,
                                    "unit_price": e_price,
                                    "sort_order": idx_e + 1,
                                    "is_pre_registered": True,
                                    "is_matched": False
                                }
                                pre_reg_count += 1

                            # 단가표에 기재된 순서 또는 앞 번호순으로 active_product_order 재정렬! ⭐
                            def get_paste_sort_key(item_tuple):
                                kp = f"{item_tuple[0]}||{item_tuple[1]}"
                                if kp in saved_pm and "sort_order" in saved_pm[kp]:
                                    return (0, int(saved_pm[kp]["sort_order"]))
                                return (1, 999999)

                            sorted_paste_items = sorted(unique_items, key=get_paste_sort_key)
                            new_paste_order = [f"{p}||{o}" for p, o in sorted_paste_items]
                            for k in saved_pm.keys():
                                if k.startswith("__official__") and k not in new_paste_order:
                                    new_paste_order.append(k)
                            st.session_state.active_product_order = new_paste_order

                            # 등록된 셀러 템플릿 및 브랜드 마스터 영구 업데이트
                            active_brand = st.session_state.get("active_brand", "호무로")
                            active_seller_tmpl = st.session_state.get("active_seller_tmpl_key")
                            active_seller_name = st.session_state.get("active_seller_name", active_seller_tmpl or "")
                            all_mappings = load_seller_mappings()
                            if active_seller_tmpl and active_seller_tmpl in all_mappings:
                                save_seller_mapping(active_seller_tmpl, active_seller_name, active_brand, user_col_map, saved_pm, product_order=new_paste_order)
                            update_brand_product_master_from_pm(active_brand, saved_pm)
                            st.session_state.active_product_mapping = saved_pm

                            for k in list(st.session_state.keys()):
                                if k.startswith("brand_price_") or k.startswith("brand_p_") or k.startswith("brand_o_"):
                                    del st.session_state[k]

                            res_msgs = []
                            if matched_count > 0:
                                res_msgs.append(f"발주서 품목 **{matched_count}개 자동 매칭**")
                            if pre_reg_count > 0:
                                res_msgs.append(f"발주서 외 공식 품목 **{pre_reg_count}개 사전등록 (⏳ 매칭대기 목록)**")
                            if res_msgs:
                                st.success(f"🎉 텍스트에서 {' 및 '.join(res_msgs)} 완료되어 공식 마스터에 영구 등록되었습니다!")
                            else:
                                st.warning("⚠️ 텍스트 안의 상품명 또는 단가 형식을 확인해주세요.")
                            st.rerun()

                # ----------------- 탭 3: 공식 품목 직접 1건 사전등록 -----------------
                with tab_manual:
                    st.markdown(f"**[{active_brand}] 공식 품목을 1건씩 직접 사전등록합니다.**")
                    st.caption("발주서에 아직 없는 품목이라도 미리 등록해 두시면, 공식 품목 마스터에 영구 보존되고 향후 발주서 업로드 시 자동 매칭 및 드롭다운 선택이 가능해집니다.")
                    man_c1, man_c2, man_c3 = st.columns([2.2, 1.4, 1.4])
                    with man_c1:
                        new_off_name = st.text_input("🏷️ 공식 품목명", placeholder="예: 올리브유 참치 80g", key="input_man_off_name")
                    with man_c2:
                        new_off_opt = st.text_input("🔢 옵션/규격/코드 (선택)", placeholder="예: 80g 또는 104012", key="input_man_off_opt")
                    with man_c3:
                        new_off_price = st.number_input("💰 본사 판매단가 (원)", min_value=0, step=1000, value=0, key="input_man_off_price")

                    if st.button("➕ 공식 품목 마스터에 영구 등록하기", type="primary", key="btn_add_manual_official_prod"):
                        if new_off_name.strip():
                            clean_oname = new_off_name.strip()
                            clean_oopt = new_off_opt.strip()
                            k_off = f"__official__{clean_oname}||{clean_oopt}"
                            saved_pm[k_off] = {
                                "brand_product": clean_oname,
                                "brand_option": clean_oopt,
                                "code": clean_oopt,
                                "unit_price": int(new_off_price),
                                "is_pre_registered": True,
                                "is_matched": False,
                                "sort_order": len(saved_pm) + 1
                            }
                            active_brand = st.session_state.get("active_brand", "호무로")
                            update_brand_product_master_from_pm(active_brand, saved_pm)
                            st.session_state.active_product_mapping = saved_pm
                            st.success(f"🎉 '[{clean_oname}]' ({int(new_off_price):,}원) 공식 품목이 마스터에 성공적으로 등록되었습니다!")
                            st.rerun()
                        else:
                            st.warning("공식 품목명을 입력해주세요.")

            # ----------------- 품목 데이터 구성 및 매칭 상태 집계 -----------------
            active_product_order = st.session_state.get("active_product_order") or seller_map_data.get("product_order", [])

            def get_sort_key(item_tuple):
                raw_p, raw_o = item_tuple
                k_pair = f"{raw_p}||{raw_o}"
                info = saved_pm.get(k_pair, {})
                bp = str(info.get("brand_product", "")).strip()

                # 1순위: 활성화된 명시적 순서 리스트
                if active_product_order and k_pair in active_product_order:
                    return (0, active_product_order.index(k_pair), 0)

                # 2순위: saved_pm에 sort_order가 지정되어 있으면 적용
                if "sort_order" in info and isinstance(info["sort_order"], (int, float)):
                    return (1, int(info["sort_order"]), 0)

                # 3순위: 브랜드 공식 상품명 또는 셀러 원본 상품명 앞의 숫자(1., 2., 3...) 추출
                num = extract_leading_number(bp) if bp else None
                if num is None:
                    num = extract_leading_number(raw_p)
                if num is not None:
                    return (2, num, 0)

                # 4순위: 그 외 원래 순서
                return (3, 999999, 0)

            unique_items = sorted(unique_items, key=get_sort_key)

            table_rows = []
            unmapped_count = 0
            mapped_count = 0

            for idx, (raw_p, raw_o) in enumerate(unique_items):
                k_pair = f"{raw_p}||{raw_o}"
                info = saved_pm.get(k_pair, {})
                bp = str(info.get("brand_product", "")).strip()
                is_keep = bool(info.get("keep_original", False) or st.session_state.get("step3_keep_original_mode", False))
                # 셀러 원본 상품명이 그대로 들어가 있거나 셀러명(에이핏 등)이 포함된 경우 정리 (단, 변경없이 진행 모드일 때는 보존)
                if not is_keep:
                    if bp == raw_p or "에이핏" in bp:
                        if k_pair in known_brand_pool and "에이핏" not in known_brand_pool[k_pair].get("brand_product", ""):
                            bp = known_brand_pool[k_pair].get("brand_product", "")
                        else:
                            bp = ""
                bo = str(info.get("brand_option", "")).strip()
                up = int(info.get("unit_price", 0))
                
                # 체크박스 확인 또는 (브랜드 상품명이 있고 0원 초과 단가)이면 매칭 완료 판정
                is_chk = bool(bp and (info.get("is_matched", False) or up > 0 or is_keep))
                if is_keep and bp:
                    is_done = "🔵 변경없음"
                elif is_chk:
                    is_done = "✅ 매칭완료"
                else:
                    is_done = "⚠️ 미매칭"
                if is_done in ["✅ 매칭완료", "🔵 변경없음"]:
                    mapped_count += 1
                else:
                    unmapped_count += 1

                table_rows.append({
                    "순번": idx + 1,
                    "상태": is_done,
                    "셀러 원본 상품명": raw_p,
                    "셀러 옵션(코드)": raw_o if raw_o else "(없음)",
                    "변환": "➔",
                    "브랜드 공식 상품명": bp,
                    "브랜드 공식 옵션": bo if bo else "(없음)",
                    "본사 판매단가": up,
                    "매칭 완료": is_chk
                })

            # ----------------- 사전등록 공식 품목(⏳ 매칭대기) 병합 -----------------
            already_matched_official = set()
            for r in table_rows:
                if r["상태"] == "✅ 매칭완료" and r["브랜드 공식 상품명"]:
                    already_matched_official.add(f"{r['브랜드 공식 상품명']}||{r['브랜드 공식 옵션']}")
                    already_matched_official.add(r['브랜드 공식 상품명'])

            pending_official_items = []
            seen_pending_keys = set()

            # 1. saved_pm 에서 사전등록 품목 탐색
            for k, v in saved_pm.items():
                if k.startswith("__official__") or v.get("is_pre_registered"):
                    bp = str(v.get("brand_product", "")).strip()
                    bo = str(v.get("brand_option", "")).strip()
                    if not bp or "에이핏" in bp:
                        continue
                    if bp in already_matched_official or f"{bp}||{bo}" in already_matched_official:
                        continue
                    p_key = f"{bp}||{bo}"
                    if p_key not in seen_pending_keys:
                        seen_pending_keys.add(p_key)
                        pending_official_items.append({
                            "brand_product": bp,
                            "brand_option": bo,
                            "code": str(v.get("code", "")).strip(),
                            "unit_price": int(v.get("unit_price", 0)),
                            "sort_order": v.get("sort_order", 999999)
                        })

            # 2. brand_product_master.json 에서도 탐색하여 누락 방지
            master_bpm = load_brand_product_master().get(active_brand, {})
            for k, v in master_bpm.items():
                bp = str(v.get("brand_product", "")).strip()
                bo = str(v.get("brand_option", "")).strip()
                if not bp or "에이핏" in bp:
                    continue
                if bp in already_matched_official or f"{bp}||{bo}" in already_matched_official:
                    continue
                p_key = f"{bp}||{bo}"
                if p_key not in seen_pending_keys:
                    seen_pending_keys.add(p_key)
                    pending_official_items.append({
                        "brand_product": bp,
                        "brand_option": bo,
                        "code": str(v.get("code", "")).strip(),
                        "unit_price": int(v.get("unit_price", 0)),
                        "sort_order": 999999
                    })

            pending_count = 0
            for p_item in pending_official_items:
                pending_count += 1
                cur_idx = len(table_rows) + 1
                bp = p_item["brand_product"]
                bo = p_item["brand_option"]
                code = p_item["code"]
                up = p_item["unit_price"]
                disp_opt = bo if bo else (code if code else "(공식목록)")
                
                table_rows.append({
                    "순번": cur_idx,
                    "상태": "⏳ 매칭대기",
                    "셀러 원본 상품명": "(사전등록 공식품목 / 발주서 미포함)",
                    "셀러 옵션(코드)": disp_opt,
                    "변환": "➔",
                    "브랜드 공식 상품명": bp,
                    "브랜드 공식 옵션": bo if bo else "(없음)",
                    "본사 판매단가": up,
                    "매칭 완료": False
                })

            # 전체 품목 행 순서를 active_product_order 및 sort_order 기준으로 정확히 정렬
            def get_final_table_row_order(r: dict) -> tuple:
                raw_p = str(r["셀러 원본 상품명"]).strip()
                raw_o = str(r["셀러 옵션(코드)"]).strip()
                if raw_o in ["(없음)", "(공식목록)", "(발주서 미포함)"]:
                    raw_o = ""
                bp = str(r.get("브랜드 공식 상품명", "")).strip()
                bo = str(r.get("브랜드 공식 옵션", "")).strip()
                if bo in ["(없음)", "(발주서 미포함)"]:
                    bo = ""
                is_pre = (raw_p == "(사전등록 공식품목 / 발주서 미포함)") or (str(r.get("상태", "")).startswith("⏳"))
                if is_pre:
                    k_pair = f"__official__{bp}||{bo}"
                else:
                    k_pair = f"{raw_p}||{raw_o}"

                if active_product_order and k_pair in active_product_order:
                    return (0, active_product_order.index(k_pair), bp)
                info = saved_pm.get(k_pair, {})
                if "sort_order" in info and isinstance(info["sort_order"], (int, float)):
                    return (1, int(info["sort_order"]), bp)
                num = extract_leading_number(bp) if bp else None
                if num is None:
                    num = extract_leading_number(raw_p)
                if num is not None:
                    return (2, num, bp)
                return (3, 999999, bp)

            table_rows.sort(key=get_final_table_row_order)
            for idx, r in enumerate(table_rows):
                r["순번"] = idx + 1

            match_table_df = pd.DataFrame(table_rows)

            # 상단 통계 카드 (발주서 현황 + 사전등록 공식품목 현황)
            sm_c1, sm_c2, sm_c3, sm_c4 = st.columns(4)
            sm_c1.metric("발주서 품목 종류", f"{len(unique_items)}개")
            sm_c2.metric("✅ 매칭 완료", f"{mapped_count}개")
            sm_c3.metric("⚠️ 미매칭 (발주서)", f"{unmapped_count}개")
            sm_c4.metric("⏳ 매칭대기 (사전등록)", f"{pending_count}개")

            # 빠른 일괄 채우기 액션 바
            if unmapped_count > 0:
                act_col1, act_col2 = st.columns([2.5, 2.5])
                with act_col1:
                    st.warning(f"⚠️ **발주서 미매칭 {unmapped_count}건** (표에서 브랜드 공식 상품명을 선택하거나 아래 1초 자동 완성을 클릭하세요)")
                with act_col2:
                    if st.button(f"⚡ [{active_brand}] 공식 매핑 1초 자동 완성", type="primary", use_container_width=True, key="btn_quick_brand_match_banner"):
                        st.session_state["step3_keep_original_mode"] = False
                        if known_brand_pool:
                            matched_dict, m_cnt = auto_match_products_against_brand_pool(unique_items, known_brand_pool)
                            if matched_dict:
                                saved_pm.update(matched_dict)
                                st.session_state.active_product_mapping = saved_pm
                                active_seller_tmpl = st.session_state.get("active_seller_tmpl_key")
                                all_mappings = load_seller_mappings()
                                if active_seller_tmpl and active_seller_tmpl in all_mappings:
                                    save_seller_mapping(active_seller_tmpl, active_seller_name, active_brand, user_col_map, saved_pm)
                                update_brand_product_master_from_pm(active_brand, saved_pm)
                                st.success(f"🎉 총 {m_cnt}개 품목이 [{active_brand}] 공식 품목명 및 판매단가로 자동 매칭되었습니다!")
                                st.rerun()
                            else:
                                st.warning("일치하는 공식 상품명을 찾지 못했습니다.")
                        else:
                            st.warning(f"아직 [{active_brand}]에 등록된 공식 상품 매핑 데이터가 없습니다.")
            elif pending_count > 0:
                st.info(f"💡 발주서 품목 매칭 완료! 사전등록된 공식 품목 **{pending_count}건이 '⏳ 매칭대기'** 상태로 보관 중입니다.")
            else:
                st.success("🎉 **모든 품목의 브랜드 매칭과 판매단가(금액) 입력이 100% 완료되었습니다!**")

            st.markdown("---")
            st.markdown("#### 엑셀 표에서 마우스 클릭으로 판매단가 & 브랜드명 직접 수정 (인터랙티브 에디터)")
            st.caption("💡 엑셀처럼 아래 표의 **'본사 판매단가'** 또는 **'브랜드 공식 상품명/옵션'** 셀을 마우스로 직접 클릭(더블클릭)하여 숫자를 입력/수정할 수 있습니다. 수정한 후 표 아래의 **[💾 템플릿 저장]** 버튼을 누르면 안전하게 저장됩니다.")

            # 브랜드 공식 상품명 후보 목록 수집 (등록된 공식 상품명 풀에서만 순수하게 수집)
            brand_opts_set = set()
            for k, inf in known_brand_pool.items():
                if inf.get("keep_original"):
                    continue
                bp_name = str(inf.get("brand_product", "")).strip()
                if bp_name and "에이핏" not in bp_name:
                    brand_opts_set.add(bp_name)

            master_bpm = load_brand_product_master().get(active_brand, {})
            for k, inf in master_bpm.items():
                if inf.get("keep_original"):
                    continue
                bp_name = str(inf.get("brand_product", "")).strip()
                if bp_name and "에이핏" not in bp_name:
                    brand_opts_set.add(bp_name)

            for k, inf in saved_pm.items():
                if inf.get("keep_original"):
                    continue
                bp_name = str(inf.get("brand_product", "")).strip()
                raw_k = k.split("||")[0].strip()
                if bp_name and bp_name != raw_k and "에이핏" not in bp_name:
                    brand_opts_set.add(bp_name)

            for val in match_table_df["브랜드 공식 상품명"].tolist():
                v_str = str(val).strip() if val else ""
                if v_str and "에이핏" not in v_str and not any(v_str == u_p for u_p, _ in unique_items):
                    brand_opts_set.add(v_str)

            # 드롭다운 공식목록 선택 시 순서가 뒤죽박죽되지 않도록 지정순번/자연번호 기준 고정 정렬 ⭐
            def get_brand_option_sort_key(bp_name: str) -> tuple:
                bp_clean = str(bp_name).strip()
                # 1. 사용자가 지정한 활성 출력순번(active_product_order) 최우선
                for s_idx, k_p in enumerate(active_product_order or []):
                    info = saved_pm.get(k_p, {})
                    if str(info.get("brand_product", "")).strip() == bp_clean:
                        return (0, s_idx, bp_clean)
                # 2. sort_order 지정 순번
                for k_p, info in saved_pm.items():
                    if str(info.get("brand_product", "")).strip() == bp_clean and "sort_order" in info:
                        return (1, int(info["sort_order"]), bp_clean)
                for k_p, info in known_brand_pool.items():
                    if str(info.get("brand_product", "")).strip() == bp_clean and "sort_order" in info:
                        return (1, int(info["sort_order"]), bp_clean)
                # 3. 상품명 앞의 숫자(1., 2., 10.) 자연 정렬
                num = extract_leading_number(bp_clean)
                if num is not None:
                    return (2, num, bp_clean)
                # 4. 장바구니/사은품/증정품 등 0원성 비매품은 항상 맨 뒤로
                if any(kw in bp_clean for kw in ["장바구니", "사은품", "증정", "미니어처"]):
                    return (4, 0, bp_clean)
                # 5. 그 외 일반 품목은 안정적인 가나다순 고정
                return (3, 0, bp_clean)

            clean_opts = sorted([x for x in brand_opts_set if x], key=get_brand_option_sort_key)
            brand_product_options = [""] + clean_opts

            curr_bp_mode = st.session_state.get("bp_input_mode_radio", "🔽 공식 목록 선택 (클릭 시 리스트 열림 ⭐)")
            if curr_bp_mode.startswith("🔽"):
                bp_col_config = st.column_config.SelectboxColumn(
                    "공식 상품명 ✏️",
                    options=brand_product_options,
                    required=False,
                    width="medium",
                    help="클릭하면 등록된 공식 상품명 목록이 열리며, 키보드로 검색하여 선택할 수도 있습니다."
                )
            else:
                bp_col_config = st.column_config.TextColumn(
                    "공식 상품명 ✏️",
                    disabled=False,
                    width="medium",
                    help="키보드로 원하는 상품명을 직접 자유롭게 타이핑하여 입력합니다."
                )

            # 컬럼 설정 (출력순번 직접 수정 가능 & 타이트한 열 너비 자동 맞춤)
            col_cfg = {
                "순번": st.column_config.NumberColumn(
                    "출력순번 ✏️",
                    format="%d",
                    min_value=1,
                    step=1,
                    disabled=False,
                    width="small",
                    help="5단계 수량표 및 최종 출력 순서를 원하는 번호(1, 2, 3...)로 직접 지정할 수 있습니다. 숫자를 바꾼 후 아래 [템플릿 저장]을 누르면 이 순서대로 자동 정렬됩니다."
                ),
                "상태": st.column_config.TextColumn("상태", disabled=True, width="small"),
                "셀러 원본 상품명": st.column_config.TextColumn("셀러 상품명", disabled=True, width="medium"),
                "셀러 옵션(코드)": st.column_config.TextColumn("셀러 옵션", disabled=True, width="small"),
                "변환": st.column_config.TextColumn("➔", disabled=True, width="small"),
                "브랜드 공식 상품명": bp_col_config,
                "브랜드 공식 옵션": st.column_config.TextColumn("공식 옵션 ✏️", disabled=False, width="small"),
                "본사 판매단가": st.column_config.NumberColumn(
                    "판매단가 ✏️",
                    min_value=0,
                    step=1000,
                    format="%d",
                    disabled=False,
                    width="small"
                ),
                "매칭 완료": st.column_config.CheckboxColumn(
                    "매칭완료 ✔️",
                    help="체크하면 단가가 0원이거나 사은품(장바구니 등)인 품목도 '매칭 완료'로 즉시 처리됩니다.",
                    default=True,
                    disabled=False,
                    width="small"
                )
            }

            def save_edited_products(df_edited: pd.DataFrame, custom_tmpl_name: str = "", is_auto_save: bool = False) -> None:
                """사용자가 에디터에서 수정한 판매단가, 공식 상품명 및 출력순번(숫자)을 영구 저장합니다."""
                active_brand = st.session_state.get("active_brand", "호무로")
                active_seller_tmpl = st.session_state.get("active_seller_tmpl_key")
                active_seller_name = st.session_state.get("active_seller_name", active_seller_tmpl or "")

                # 1. '순번' 열을 숫자로 변환 후 오름차순 정렬 (사용자가 입력한 숫자에 따라 순서 자동 재배치 ⭐)
                df_to_save = df_edited.copy()
                if "순번" in df_to_save.columns:
                    df_to_save["순번"] = pd.to_numeric(df_to_save["순번"], errors="coerce").fillna(999999)
                    df_to_save = df_to_save.sort_values(by="순번").reset_index(drop=True)

                count_saved = 0
                new_order_keys = []
                for rank_idx, r in df_to_save.iterrows():
                    raw_p = str(r["셀러 원본 상품명"]).strip()
                    raw_o = str(r["셀러 옵션(코드)"]).strip()
                    if raw_o in ["(없음)", "(공식목록)", "(발주서 미포함)"]:
                        raw_o = ""
                    bp = str(r.get("브랜드 공식 상품명", "")).strip()
                    is_keep = bool(r.get("keep_original", False) or st.session_state.get("step3_keep_original_mode", False) or (bp and bp == raw_p))
                    if not is_keep:
                        if bp == raw_p or "에이핏" in bp:
                            bp = ""
                    bo = str(r.get("브랜드 공식 옵션", "")).strip()
                    if bo in ["(없음)", "(발주서 미포함)"]:
                        bo = ""
                    try:
                        up_val = int(float(str(r.get("본사 판매단가", 0)).replace(",", "").replace("원", "")))
                    except Exception:
                        up_val = 0

                    is_pre = (raw_p == "(사전등록 공식품목 / 발주서 미포함)") or (str(r.get("상태", "")).startswith("⏳"))
                    if is_pre:
                        if not bp:
                            continue
                        k_pair = f"__official__{bp}||{bo}"
                        is_matched_final = False
                    else:
                        k_pair = f"{raw_p}||{raw_o}"
                        # 만약 단가가 0원인데 공식 상품명이 선택되었고, 브랜드 풀에 해당 상품의 단가가 등록되어 있다면 자동 보충
                        if up_val == 0 and bp and not is_keep:
                            for p_k, p_inf in known_brand_pool.items():
                                if p_inf.get("brand_product") == bp and int(p_inf.get("unit_price", 0)) > 0:
                                    up_val = int(p_inf["unit_price"])
                                    break
                        is_chk = bool(r.get("매칭 완료", False))
                        is_matched_final = bool(bp and (is_chk or up_val > 0 or is_keep))

                    new_order_keys.append(k_pair)
                    # 사용자가 에디터에서 직접 입력한 순번 숫자를 우선 적용
                    cur_sort_num = int(r["순번"]) if "순번" in r and pd.notna(r["순번"]) and int(r["순번"]) < 999999 else (rank_idx + 1)
                    saved_pm[k_pair] = {
                        "brand_product": bp,
                        "brand_option": bo,
                        "code": bo if bo else str(r.get("셀러 옵션(코드)", "")).replace("(공식목록)", "").strip(),
                        "unit_price": up_val,
                        "is_matched": is_matched_final,
                        "is_pre_registered": is_pre,
                        "keep_original": is_keep,
                        "sort_order": cur_sort_num
                    }
                    count_saved += 1

                tmpl_to_save = (custom_tmpl_name or "").strip()
                if not tmpl_to_save:
                    tmpl_to_save = active_seller_tmpl or active_tmpl_key or (seller_file_name if seller_file_name else "기본셀러템플릿")

                # 전체 품목 순서 재구성: 모든 품목을 sort_order(숫자) 기준 오름차순으로 안정적으로 정렬
                all_keys_in_scope = []
                for r_k in new_order_keys:
                    if r_k not in all_keys_in_scope:
                        all_keys_in_scope.append(r_k)
                cur_existing_order = st.session_state.get("active_product_order") or []
                for old_k in cur_existing_order:
                    if old_k not in all_keys_in_scope:
                        all_keys_in_scope.append(old_k)
                for pm_k in saved_pm.keys():
                    if pm_k not in all_keys_in_scope:
                        all_keys_in_scope.append(pm_k)

                def get_key_final_order(k: str) -> tuple:
                    inf = saved_pm.get(k, {})
                    if "sort_order" in inf and isinstance(inf["sort_order"], (int, float)):
                        return (0, int(inf["sort_order"]))
                    if k in cur_existing_order:
                        return (1, cur_existing_order.index(k))
                    return (2, 999999)

                full_order = sorted(all_keys_in_scope, key=get_key_final_order)
                # full_order 순서대로 saved_pm의 sort_order를 1부터 순차 재부여 (1, 2, 3...)
                for f_idx, f_k in enumerate(full_order):
                    if f_k in saved_pm:
                        saved_pm[f_k]["sort_order"] = f_idx + 1

                # 셀러 템플릿 JSON에 영구 저장 (새 템플릿 등록 및 기존 템플릿 갱신, product_order 포함!)
                save_seller_mapping(tmpl_to_save, tmpl_to_save, active_brand, user_col_map, saved_pm, product_order=full_order)
                update_brand_product_master_from_pm(active_brand, saved_pm)

                st.session_state.active_seller_tmpl_key = tmpl_to_save
                st.session_state.active_seller_name = tmpl_to_save
                st.session_state.active_product_mapping = saved_pm
                st.session_state.active_product_order = full_order
                if not is_auto_save:
                    st.success(f"✨ '{tmpl_to_save}' 템플릿에 총 {count_saved}개 품목의 판매단가, 매칭 상태 및 출력 순서가 영구 저장되었습니다!")
                    st.rerun()

            # 검색 필터 & 공식 상품명 입력 모드 토글
            f_col1, f_col2 = st.columns([2.3, 1.7])
            with f_col1:
                filter_search = st.text_input("🔍 품목명 검색 (상품명, 코드 등으로 빠른 찾기)", "", key="match_search_filter")
            with f_col2:
                bp_input_mode = st.radio(
                    "공식 상품명 입력 방식",
                    ["🔽 공식 목록 선택 (클릭 시 리스트 열림 ⭐)", "✏️ 자유 텍스트 직접 타이핑"],
                    horizontal=True,
                    key="bp_input_mode_radio",
                    help="'공식 목록 선택' 시 표에서 셀을 클릭하면 기존에 매핑된 공식 상품명 리스트가 열립니다. 신규 상품명을 자유롭게 직접 치려면 '직접 타이핑'을 선택하세요."
                )

            def apply_filter(df_src):
                if not filter_search.strip():
                    return df_src
                kw = filter_search.strip().lower()
                mask = (
                    df_src["셀러 원본 상품명"].astype(str).str.lower().str.contains(kw) |
                    df_src["셀러 옵션(코드)"].astype(str).str.lower().str.contains(kw) |
                    df_src["브랜드 공식 상품명"].astype(str).str.lower().str.contains(kw) |
                    df_src["브랜드 공식 옵션"].astype(str).str.lower().str.contains(kw)
                )
            # ----------------- 5단계 수량표 출력 순서 안내 (간결한 안내) -----------------
            st.caption("💡 **수량표 출력 순서 지정**: 아래 표의 **'출력순번 ✏️'** 열 숫자를 클릭하여 원하는 순서 번호(1, 2, 3...)를 직접 입력하고 **[💾 템플릿 저장]**을 누르면, 5단계 수량표에 해당 순서대로 정렬되어 출력됩니다.")

            tab_view_unmap, tab_view_all, tab_view_map = st.tabs([
                f"⚠️ 1. 미매칭 & ⏳ 매칭대기 품목 (집중 입력/확인 🎯) ({unmapped_count + pending_count}건)",
                f"📋 2. 전체 품목 대조표 (직접 클릭 수정 ✏️) ({len(match_table_df)}건)",
                f"✅ 3. 매칭 완료 품목 ({mapped_count}건)"
            ])

            with tab_view_unmap:
                df_unmap = match_table_df[match_table_df["상태"].isin(["⚠️ 미매칭", "⏳ 매칭대기"])].copy()
                df_unmap_filtered = apply_filter(df_unmap).copy()

                if df_unmap.empty:
                    st.success("🎉 미매칭 및 매칭대기 품목이 없습니다! 모든 품목이 매칭 완료되었습니다.")
                    st.info("💡 전체 품목의 순번 변경이나 단가 수정은 상단의 **'📋 2. 전체 품목 대조표'** 탭을 클릭하여 진행하세요.")
                elif df_unmap_filtered.empty:
                    st.info("검색 조건에 맞는 미매칭/매칭대기 품목이 없습니다.")
                else:
                    st.caption("👇 아래 표에서 **'판매단가'** 열을 직접 클릭하여 단가를 입력하세요. 단가가 0원이거나 사은품/장바구니인 경우 오른쪽 끝 **'매칭완료 ✔️'**를 체크하시면 됩니다. 사전등록된 **'⏳ 매칭대기'** 품목의 단가/품목명도 직접 수정 가능합니다.")
                    calc_h = min(500, max(180, 45 + len(df_unmap_filtered) * 38))
                    edited_unmap = st.data_editor(
                        df_unmap_filtered,
                        column_config=col_cfg,
                        hide_index=True,
                        use_container_width=True,
                        height=calc_h,
                        key="editor_unmap_view"
                    )
                    # 표에서 직접 수정한 단가/옵션/순번을 세션에 실시간 자동 보존 ⭐
                    if st.session_state.get("editor_unmap_view", {}).get("edited_rows"):
                        target_name_auto = str(st.session_state.get("active_seller_tmpl_key") or active_tmpl_key or (seller_file_name if seller_file_name else "기본셀러템플릿"))
                        save_edited_products(edited_unmap, target_name_auto, is_auto_save=True)

                    st.markdown('<div class="save-template-btn">', unsafe_allow_html=True)
                    s_c1, s_c2 = st.columns([2.5, 1.2], vertical_alignment="bottom")
                    with s_c1:
                        target_tmpl_input_1 = st.text_input(
                            "저장할 셀러 템플릿 이름",
                            value=str(st.session_state.get("active_seller_tmpl_key") or active_tmpl_key or (seller_file_name if seller_file_name else "")),
                            key="input_save_tmpl_name_tab1",
                            help="이 이름의 템플릿으로 저장되어 상단 2번 '셀러 템플릿 가져오기'에서 언제든 다시 불러올 수 있습니다."
                        )
                    with s_c2:
                        if st.button("💾 템플릿 저장", type="primary", use_container_width=True, key="btn_save_unmap_editor"):
                            save_edited_products(edited_unmap, target_tmpl_input_1)
                    st.markdown('</div>', unsafe_allow_html=True)

            with tab_view_all:
                df_all_filtered = apply_filter(match_table_df).copy()

                if df_all_filtered.empty:
                    st.info("검색 조건에 맞는 품목이 없습니다.")
                else:
                    st.caption("👇 발주서 품목과 사전등록 공식 품목(⏳ 매칭대기)의 공식 상품명과 판매단가를 한눈에 대조합니다. **'출력순번 ✏️'** 숫자를 원하는 순서(1, 2, 3...)로 직접 변경하거나 단가를 수정한 후 아래 **[💾 템플릿 저장]**을 누르면 5단계 일자별 수량표에 그대로 반영됩니다.")
                    calc_h_all = min(550, max(200, 45 + len(df_all_filtered) * 38))
                    edited_all = st.data_editor(
                        df_all_filtered,
                        column_config=col_cfg,
                        hide_index=True,
                        use_container_width=True,
                        height=calc_h_all,
                        key="editor_all_view"
                    )
                    # 표에서 직접 수정한 단가/옵션/순번을 세션에 실시간 자동 보존 ⭐
                    if st.session_state.get("editor_all_view", {}).get("edited_rows"):
                        target_name_auto = str(st.session_state.get("active_seller_tmpl_key") or active_tmpl_key or (seller_file_name if seller_file_name else "기본셀러템플릿"))
                        save_edited_products(edited_all, target_name_auto, is_auto_save=True)

                    st.markdown('<div class="save-template-btn">', unsafe_allow_html=True)
                    s_c1_all, s_c2_all = st.columns([2.5, 1.2], vertical_alignment="bottom")
                    with s_c1_all:
                        target_tmpl_input_2 = st.text_input(
                            "저장할 셀러 템플릿 이름",
                            value=str(st.session_state.get("active_seller_tmpl_key") or active_tmpl_key or (seller_file_name if seller_file_name else "")),
                            key="input_save_tmpl_name_tab2",
                            help="이 이름의 템플릿으로 저장되어 상단 2번 '셀러 템플릿 가져오기'에서 언제든 다시 불러올 수 있습니다."
                        )
                    with s_c2_all:
                        if st.button("💾 템플릿 저장", type="primary", use_container_width=True, key="btn_save_all_editor"):
                            save_edited_products(edited_all, target_tmpl_input_2)
                    st.markdown('</div>', unsafe_allow_html=True)

            with tab_view_map:
                df_map = match_table_df[match_table_df["상태"] == "✅ 매칭완료"].copy()
                df_map_filtered = apply_filter(df_map).copy()

                if df_map.empty:
                    st.warning("아직 매칭이 완료된 품목이 없습니다. 1번 탭에서 단가를 입력하거나 '매칭완료 체크'를 선택해주세요.")
                elif df_map_filtered.empty:
                    st.info("검색 조건에 맞는 매칭 완료 품목이 없습니다.")
                else:
                    st.dataframe(df_map_filtered, column_config=col_cfg, hide_index=True, use_container_width=True)

            st.session_state.active_product_mapping = saved_pm

            # Step 3 셀러 템플릿 누적 품목 전체 관리 & 마스터 엑셀 다운로드
            tmpl_display_name = str(active_seller_tmpl or active_tmpl_key or (seller_file_name if seller_file_name else "현재 템플릿")).strip()
            with st.expander(f"📚 '{tmpl_display_name}' 템플릿에 누적된 전체 매핑 품목 열람 & 관리 ({len(saved_pm)}건)", expanded=False):
                st.markdown("""
                <div style="background-color: #F8F9FA; border-left: 4px solid #1E3D59; padding: 10px 14px; border-radius: 4px; margin-bottom: 12px; font-size: 13px; color: #444; line-height: 1.6;">
                    <strong>💡 셀러 매핑 품목이 계속 누적되는 이유:</strong><br>
                    셀러 발주서 양식이 바뀌거나 회차마다 다른 옵션/행사 품목이 들어올 때마다, 다음번 작업 시 <strong>1초 자동 매칭</strong>에 활용할 수 있도록 템플릿 매핑 사전에 자동으로 누적(Merge) 기억됩니다.<br>
                    • <strong>위 3단계 대조표:</strong> 지금 업로드된 발주서에 실제로 들어있는 품목만 집중 작업하도록 표시됩니다.<br>
                    • <strong>아래 표:</strong> 이 템플릿에 영구 기억된 <strong>전체 {0}개 품목</strong> 목록이며, 더 이상 쓰지 않는 과거 품목이나 오타 품목을 직접 선택하여 삭제할 수 있습니다.
                </div>
                """.format(len(saved_pm)), unsafe_allow_html=True)

                # 누적된 전체 품목 DataFrame 구성
                cur_order_keys = set(f"{p}||{o}" for p, o in unique_items)
                accum_rows = []
                for idx_ac, (k_p, p_inf) in enumerate(saved_pm.items()):
                    is_pre = bool(p_inf.get("is_pre_registered") or str(k_p).startswith("__official__"))
                    clean_kp = str(k_p).replace("__official__", "")
                    parts = clean_kp.split("||")
                    raw_p = parts[0] if len(parts) > 0 else ""
                    raw_o = parts[1] if len(parts) > 1 else ""
                    bp = str(p_inf.get("brand_product", "")).strip()
                    bo = str(p_inf.get("brand_option", "")).strip()
                    up = int(p_inf.get("unit_price", 0))

                    if k_p in cur_order_keys:
                        src_type = "📦 이번 발주서 포함"
                    elif is_pre:
                        src_type = "⏳ 사전등록 공식품목"
                    else:
                        src_type = "💾 템플릿 보관 (과거 이력)"

                    accum_rows.append({
                        "키": k_p,
                        "순번": idx_ac + 1,
                        "구분": src_type,
                        "셀러 원본 상품명": raw_p if raw_p else "(사전등록 공식품목)",
                        "셀러 옵션(코드)": raw_o if raw_o else "(없음)",
                        "➔": "➔",
                        "브랜드 공식 상품명": bp if bp else "(미지정)",
                        "브랜드 공식 옵션": bo if bo else "(없음)",
                        "판매단가(원)": up
                    })

                if accum_rows:
                    df_accum = pd.DataFrame(accum_rows)
                    df_accum_disp = df_accum.drop(columns=["키"])
                    st.dataframe(df_accum_disp, hide_index=True, use_container_width=True)

                    # 1. 불필요한 매핑 품목 개별 선택 삭제
                    st.markdown("##### 🗑️ 불필요한 누적 매핑 품목 개별 삭제")
                    del_c1, del_c2 = st.columns([3, 1], vertical_alignment="bottom")
                    with del_c1:
                        del_target_options = [r["키"] for r in accum_rows]
                        def fmt_del_opt(k_val: str) -> str:
                            inf = saved_pm.get(k_val, {})
                            bp_show = inf.get("brand_product") or k_val.replace("__official__", "")
                            clean_disp = k_val.replace("__official__", "")
                            return f"[{clean_disp}] ➔ {bp_show} ({inf.get('unit_price', 0):,}원)"

                        del_keys_selected = st.multiselect(
                            "삭제할 매핑 품목 선택 (과거 테스트 품목이나 오타 품목)",
                            del_target_options,
                            format_func=fmt_del_opt,
                            key="ms_delete_accum_pm_items"
                        )
                    with del_c2:
                        if st.button("🗑️ 선택 품목 매핑 삭제", type="secondary", use_container_width=True, key="btn_del_selected_accum_pm"):
                            if del_keys_selected:
                                for dk in del_keys_selected:
                                    if dk in saved_pm:
                                        del saved_pm[dk]
                                st.session_state.active_product_mapping = saved_pm
                                tmpl_save = active_seller_tmpl or active_tmpl_key or (seller_file_name if seller_file_name else "기본셀러템플릿")
                                cur_order = st.session_state.get("active_product_order") or []
                                cur_order = [ok for ok in cur_order if ok not in del_keys_selected]
                                st.session_state.active_product_order = cur_order
                                save_seller_mapping(tmpl_save, active_seller_name or tmpl_save, active_brand, user_col_map, saved_pm, product_order=cur_order)
                                st.success(f"🎉 총 {len(del_keys_selected)}개 품목 매핑이 안전하게 삭제되었습니다!")
                                st.rerun()
                            else:
                                st.warning("삭제할 품목을 선택해주세요.")

                    # 2. 이번 발주서에 없는 과거 이력 품목 일괄 정리
                    st.markdown("---")
                    clean_c1, clean_c2 = st.columns([2.8, 1.2], vertical_alignment="center")
                    with clean_c1:
                        st.caption("💡 **팁:** 현재 발주서에 들어있는 품목들만 남기고, 이전 작업에서 누적된 과거 이력 품목을 한 번에 정리하시려면 우측 버튼을 누르세요.")
                    with clean_c2:
                        if st.button("🧹 과거 이력 일괄 정리", type="secondary", use_container_width=True, key="btn_clean_unused_pm"):
                            removed_cnt = 0
                            for k_p in list(saved_pm.keys()):
                                if k_p not in cur_order_keys and not str(k_p).startswith("__official__") and not saved_pm[k_p].get("is_pre_registered"):
                                    del saved_pm[k_p]
                                    removed_cnt += 1
                            if removed_cnt > 0:
                                st.session_state.active_product_mapping = saved_pm
                                tmpl_save = active_seller_tmpl or active_tmpl_key or (seller_file_name if seller_file_name else "기본셀러템플릿")
                                cur_order = st.session_state.get("active_product_order") or []
                                cur_order = [ok for ok in cur_order if ok in saved_pm]
                                st.session_state.active_product_order = cur_order
                                save_seller_mapping(tmpl_save, active_seller_name or tmpl_save, active_brand, user_col_map, saved_pm, product_order=cur_order)
                                st.success(f"🎉 과거 발주서 이력 {removed_cnt}건이 깔끔하게 정리되었습니다!")
                                st.rerun()
                            else:
                                st.info("정리할 과거 이력 품목이 없습니다.")

                # 3. 마스터 매핑표 엑셀 다운로드
                st.markdown("---")
                st.caption("현재 저장된 품목명 대조 및 본사 판매단가 마스터(사전등록 공식 품목 포함)를 엑셀(.xlsx) 파일로 내보냅니다.")
                if saved_pm or known_brand_pool:
                    pm_excel_bytes = export_product_mapping_excel(saved_pm, brand_master_dict=known_brand_pool)
                    st.download_button(
                        label="📥 품목·판매단가 마스터 매핑표 다운로드 (.xlsx)",
                        data=pm_excel_bytes,
                        file_name=f"품목단가_마스터매핑_{active_brand}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.info("아직 저장된 품목 매핑 데이터가 없습니다.")

            st.markdown("""
            <div class="notion-callout notion-callout-blue">
                <span style="font-size: 20px;">👉</span>
                <div>
                    <strong>품목 매칭 및 본사 판매단가 확인이 완료되었으면 4단계로 이동하세요.</strong><br>
                    위 표에서 <strong>'💾 템플릿 저장'</strong>을 완료한 후, <strong>'👉 4단계 이벤트 & 사은품 맞춤 설정으로 이동'</strong>을 눌러 프로모션 설정을 진행하세요.
                </div>
            </div>
            """, unsafe_allow_html=True)

            if st.button("👉 4단계 이벤트 & 사은품 맞춤 설정으로 이동", type="primary", use_container_width=True, key="btn_goto_step4_from3"):
                # 에디터에 변경된 순번/단가가 있으면 4단계 이동 전에 자동 저장하여 유실 방지
                try:
                    active_seller_tmpl = st.session_state.get("active_seller_tmpl_key")
                    target_tmpl_name = active_seller_tmpl or active_tmpl_key or (seller_file_name if seller_file_name else "기본셀러템플릿")
                    if 'edited_all' in locals() and not df_all_filtered.empty:
                        save_edited_products(edited_all, target_tmpl_name, is_auto_save=True)
                    elif 'edited_unmap' in locals() and not df_unmap_filtered.empty:
                        save_edited_products(edited_unmap, target_tmpl_name, is_auto_save=True)
                except Exception:
                    pass
                st.session_state["active_step_idx"] = 3
                st.rerun()
    else:
        st.info("👆 먼저 2단계에서 셀러 발주서 파일을 업로드하고 헤더를 매칭해주세요.")
        if st.button("👈 2단계 셀러 발주서 업로드로 이동", type="secondary", key="btn_goto_step2_from3_empty"):
            st.session_state["active_step_idx"] = 1
            st.rerun()

# =========================================================================
# STEP 4: 이벤트 & 사은품 커스텀 설정
# =========================================================================
elif cur_step_idx == 3:
    st.markdown("""
    <div class="notion-step-banner">🎁 4단계. 이벤트 & 사은품 맞춤 설정</div>
    <div class="notion-step-desc-big">선착순, 금액달성, 특정품목 증정 등 각 프로모션별 반영 방식과 대상을 자유롭게 커스텀하세요. 행사가 없을 때는 원클릭으로 건너뛸 수 있습니다.</div>
    """, unsafe_allow_html=True)

    # 1. 상단 원클릭 빠른 액션 바 (행사 없음 vs 사은품 적용하고 5단계 이동)
    st.markdown("""
    <div class="notion-callout notion-callout-yellow">
        <span style="font-size: 24px;">💡</span>
        <div>
            <strong style="font-size: 15px;">사은품 행사가 없으신가요? 아니면 설정을 완료하셨나요?</strong><br>
            행사가 없는 일반 발주는 <strong>'이벤트 없음'</strong>을 누르시고, 사은품 행사를 진행하실 때는 아래 설정을 체크한 뒤 <strong>'사은품 설정 완료 & 5단계 이동'</strong> 버튼을 눌러주세요.
        </div>
    </div>
    """, unsafe_allow_html=True)

    event_templates = load_event_templates()
    saved_ev_keys = list(event_templates.keys())

    # ---------------- 1. 상단: 이벤트 템플릿 선택 & 관리 ----------------
    st.markdown("""
    <div style="font-size: 15px; font-weight: 700; color: #37352F; margin: 12px 0 4px 0;">
        📑 이벤트 템플릿 선택 & 관리 (저장된 목록)
    </div>
    """, unsafe_allow_html=True)
    st.caption("저장된 이벤트 템플릿 목록에서 원하는 설정을 선택하면 선착순/금액달성/특정품목/사은품 규칙이 1초 만에 자동 로드됩니다.")

    col_ev_select, col_ev_pin, col_ev_del = st.columns([2.5, 0.7, 1.1], vertical_alignment="bottom")
    with col_ev_select:
        if "pending_event_tmpl_choice" in st.session_state:
            st.session_state["sb_event_tmpl_choice"] = st.session_state.pop("pending_event_tmpl_choice")

        ev_options = ["+ 새 이벤트 템플릿 설정 (직접 구성)"] + saved_ev_keys
        curr_ev_choice = st.session_state.get("sb_event_tmpl_choice", ev_options[0])
        if curr_ev_choice not in ev_options:
            curr_ev_choice = ev_options[0]
        def_ev_idx = ev_options.index(curr_ev_choice)

        selected_ev_tmpl = st.selectbox(
            "적용할 이벤트 템플릿 선택",
            ev_options,
            index=def_ev_idx,
            key="sb_event_tmpl_choice"
        )
    with col_ev_pin:
        if selected_ev_tmpl != "+ 새 이벤트 템플릿 설정 (직접 구성)":
            is_pinned_curr = bool(event_templates.get(selected_ev_tmpl, {}).get("is_pinned", False))
            new_pin = st.checkbox(
                "📌 고정",
                value=is_pinned_curr,
                key=f"chk_pin_tmpl_{selected_ev_tmpl}",
                help="체크하면 아래의 '고정된 자주 쓰는 이벤트' 목록에 원클릭 바로가기 버튼으로 등록됩니다."
            )
            if new_pin != is_pinned_curr:
                toggle_event_template_pinned(selected_ev_tmpl, new_pin)
                st.rerun()
    with col_ev_del:
        if selected_ev_tmpl not in ["+ 새 이벤트 템플릿 설정 (직접 구성)", "기본 (이벤트 없음)"]:
            if st.button("🗑️ 템플릿 삭제", key="btn_del_ev_tmpl", type="secondary", use_container_width=True):
                delete_event_template(selected_ev_tmpl)
                st.warning(f"'{selected_ev_tmpl}' 이벤트 템플릿이 삭제되었습니다.")
                st.session_state["last_ev_sync_context"] = None
                st.rerun()

    # 이벤트 템플릿 선택 시 설정 동기화
    active_sel_name = st.session_state.get('active_seller_name') or extract_clean_seller_from_filename(st.session_state.get('current_seller_file', '')) or '셀러'
    active_b_name = st.session_state.get('active_brand') or '브랜드'
    default_ev_title = selected_ev_tmpl if selected_ev_tmpl != "+ 새 이벤트 템플릿 설정 (직접 구성)" else f"{active_sel_name}_{active_b_name}"

    sync_ev_context = f"ev_tmpl___{selected_ev_tmpl}"
    if st.session_state.get("last_ev_sync_context") != sync_ev_context:
        if selected_ev_tmpl in event_templates:
            loaded_cfg = event_templates[selected_ev_tmpl].get("config", {})
        else:
            loaded_cfg = default_event_config
        apply_event_template_to_session(loaded_cfg, selected_ev_tmpl, default_ev_title)
        st.session_state["last_ev_sync_context"] = sync_ev_context
        st.rerun()

    # 고정된 이벤트 템플릿만 원클릭 바로가기 버튼으로 표시 ⭐
    pinned_ev_keys = [k for k in saved_ev_keys if event_templates.get(k, {}).get("is_pinned", False)]
    if pinned_ev_keys:
        st.markdown("<div style='margin: 8px 0 4px 0;'><strong style='font-size:13px; color:#1E3D59;'>📌 고정된 이벤트 템플릿 (클릭 시 즉시 설정):</strong></div>", unsafe_allow_html=True)
        n_pins = len(pinned_ev_keys)
        for chunk_start in range(0, n_pins, 4):
            chunk_keys = pinned_ev_keys[chunk_start:chunk_start + 4]
            btn_cols = st.columns([1] * len(chunk_keys) + ([max(1, 4 - len(chunk_keys))] if len(chunk_keys) < 4 else []))
            for c_idx, pk in enumerate(chunk_keys):
                with btn_cols[c_idx]:
                    is_sel = (selected_ev_tmpl == pk)
                    btn_type = "primary" if is_sel else "secondary"
                    btn_label = f"⭐ {pk}" if is_sel else f"📌 {pk}"
                    if st.button(btn_label, key=f"btn_quick_pin_ev_{chunk_start + c_idx}_{pk}", type=btn_type, use_container_width=True, help=f"클릭하면 '{pk}' 이벤트 설정이 즉시 반영됩니다."):
                        st.session_state["pending_event_tmpl_choice"] = pk
                        st.session_state["last_ev_sync_context"] = None
                        st.rerun()

    # 템플릿 명칭 입력란 (셀러_브랜드 형식 ⭐)
    cur_ev_title = st.session_state.get("input_event_tmpl_name_widget", "")
    if selected_ev_tmpl == "+ 새 이벤트 템플릿 설정 (직접 구성)":
        if not cur_ev_title or "호무로" in cur_ev_title or cur_ev_title.endswith("_이벤트"):
            st.session_state["input_event_tmpl_name_widget"] = default_ev_title
            cur_ev_title = default_ev_title
    else:
        st.session_state["input_event_tmpl_name_widget"] = selected_ev_tmpl
        cur_ev_title = selected_ev_tmpl

    input_ev_tmpl_title = st.text_input(
        "🏷️ 이벤트 템플릿 이름 (저장할 명칭)",
        value=cur_ev_title,
        key="input_event_tmpl_name_widget",
        placeholder="예: 셀러_브랜드 (예: 공구셀러_템포)"
    )

    cfg_cur = st.session_state.get("event_config", default_event_config)

    def parse_action_rule(rule_str):
        action_type = "tag" if "행 추가 안 함" in rule_str else "add_row"
        gift_pos = rule_str
        if "배송메세지" in rule_str:
            tag_display = "배송메세지 / 배송메모 열에 텍스트 덧붙임"
        else:
            tag_display = "텍스트 표기 없음 (화면 및 엑셀 배경색 하이라이트 색인만 처리 ⭐)"
        return action_type, gift_pos, tag_display

    top_save_template_clicked = False
    top_apply_clicked = False

    # 상단 2단 액션 버튼 바 (왼쪽: 이벤트 적용하고 5단계 이동 ⭐ / 오른쪽: 이벤트 없음 건너뛰기)
    col_top_apply, col_top_skip = st.columns([1.1, 1.3])
    with col_top_apply:
        if st.button("👉 이벤트 적용하고 5단계 최종 발주서 생성", type="primary", use_container_width=True, key="btn_apply_events_top"):
            top_apply_clicked = True
    with col_top_skip:
        if st.button("⏩ 이벤트 없음 (사은품 행사 없이 바로 5단계 최종 발주서로 이동)", type="secondary", use_container_width=True, key="btn_skip_all_events"):
            st.session_state["event_skip_mode"] = True
            st.session_state["event_config_confirmed"] = True
            st.session_state["event_config"] = {
                "color_event": "#FFF2CC",
                "color_gift": "#E2EFDA",
                "first_come_first_served": {"enabled": False},
                "min_amount": {"enabled": False},
                "special_product": {"enabled": False},
                "custom_events": []
            }
            st.session_state["active_step_idx"] = 4
            st.rerun()

    # 현재 템플릿의 활성화된 이벤트 규칙 한눈에 보기 요약 배너 ⭐
    cur_fcfs = cfg_cur.get("first_come_first_served", {})
    cur_amt = cfg_cur.get("min_amount", {})
    cur_spec = cfg_cur.get("special_product", {})
    custom_events_cur_list = st.session_state.get("custom_events_list", cfg_cur.get("custom_events", []))

    active_summary_items = []
    if cur_fcfs.get("enabled"):
        active_summary_items.append(f"⏱️ <strong>선착순 {cur_fcfs.get('limit', 3)}명</strong>: {cur_fcfs.get('gift_product_name', '사은품')}")
    if cur_amt.get("enabled"):
        fcfs_extra = f" (선착순 {cur_amt.get('fcfs_limit')}명 한정)" if cur_amt.get("fcfs_enabled") else ""
        active_summary_items.append(f"💰 <strong>{cur_amt.get('min_amount', 50000):,}원 이상</strong>{fcfs_extra}: {cur_amt.get('gift_product_name', '사은품')}")
    if cur_spec.get("enabled"):
        active_summary_items.append(f"🎁 <strong>품목 '{cur_spec.get('keyword', '')}'</strong>: {cur_spec.get('gift_product_name', '사은품')}")
    for cev in custom_events_cur_list:
        if cev.get("enabled"):
            active_summary_items.append(f"✨ <strong>{cev.get('title', '이벤트')}</strong>: {cev.get('gift_product_name', '사은품')}")

    if active_summary_items:
        summary_html = " &nbsp;|&nbsp; ".join(active_summary_items)
        st.markdown(f"""
        <div style="background: #EEF8F6; border: 1px solid #C4E7E1; border-left: 4px solid #0F7B6C; padding: 10px 14px; border-radius: 6px; margin: 10px 0 14px 0; font-size: 13.5px; color: #1E3D59;">
            <div style="font-weight: 700; margin-bottom: 3px; color: #0F7B6C;">✅ 현재 불러온 템플릿에서 활성화된 이벤트 ({len(active_summary_items)}개):</div>
            <div>{summary_html}</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="background: #F8F9FA; border: 1px solid #E5E7EB; border-left: 4px solid #9E9E9E; padding: 8px 12px; border-radius: 6px; margin: 10px 0 14px 0; font-size: 13px; color: #616161;">
            💡 현재 활성화된 사은품 이벤트가 없습니다. 아래의 메뉴에서 원하는 이벤트를 체크하여 설정하세요.
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("📋 개별 프로모션 이벤트 독립 커스텀 (필요한 행사만 체크)")

    # 1. 선착순 이벤트 (활성화 시 자동 펼침)
    exp_fcfs = bool(cur_fcfs.get("enabled", False))
    with st.expander(f"⏱️ 1. 선착순 주문 사은품 이벤트 {'(✅ 활성)' if exp_fcfs else ''}", expanded=exp_fcfs):
        st.caption("주문일시가 가장 빠른 고객에게 선착순으로 사은품을 지급합니다.")
        fcfs_on = st.checkbox("선착순 사은품 적용하기", value=cur_fcfs.get("enabled", False), key="ev_fcfs_on")
        
        fc_c1, fc_c2 = st.columns([1.2, 1.2])
        with fc_c1:
            fcfs_limit = st.number_input("선착순 인정 인원(명)", min_value=1, value=cur_fcfs.get("limit", 3), step=1, key="ev_fcfs_limit")
            
            # 체크박스 기반 판정 기준 (사용자 요청 반영 ⭐)
            st.write("**선착순 판정 기준 (동일 고객 체크)**")
            saved_fcfs_targets = cur_fcfs.get("target_rule", ["수취인명", "연락처"])
            if isinstance(saved_fcfs_targets, str):
                saved_fcfs_targets = [saved_fcfs_targets]
            
            chk_c1, chk_c2, chk_c3, chk_c4 = st.columns(4)
            with chk_c1:
                fcfs_k_oid = st.checkbox("주문번호", value=any("주문번호" in str(x) for x in saved_fcfs_targets), key="fcfs_chk_oid")
            with chk_c2:
                fcfs_k_name = st.checkbox("수취인명", value=any(k in str(x) for x in saved_fcfs_targets for k in ["수취인", "수령인", "구매자", "1인"]), key="fcfs_chk_name")
            with chk_c3:
                fcfs_k_phone = st.checkbox("연락처", value=any(k in str(x) for x in saved_fcfs_targets for k in ["연락처", "전화", "1인"]), key="fcfs_chk_phone")
            with chk_c4:
                fcfs_k_addr = st.checkbox("배송주소", value=any("주소" in str(x) for x in saved_fcfs_targets), key="fcfs_chk_addr")

            fcfs_crit_list = []
            if fcfs_k_oid: fcfs_crit_list.append("주문번호")
            if fcfs_k_name: fcfs_crit_list.append("수취인명")
            if fcfs_k_phone: fcfs_crit_list.append("연락처")
            if fcfs_k_addr: fcfs_crit_list.append("주소")
            if not fcfs_crit_list:
                fcfs_crit_list = ["수취인명", "연락처"]
            
            st.caption(f"💡 체크 기준: **{' + '.join(fcfs_crit_list)}** 동일 시 1인 1회로 판정하여 중복 수령 방지")

        with fc_c2:
            # 우측 위 반영 방식 기준 설정 (사용자 요청 반영 ⭐)
            curr_fcfs_act = cur_fcfs.get("gift_position", ACTION_RULE_OPTIONS[0])
            fcfs_act_idx = 0
            for idx_a, opt_a in enumerate(ACTION_RULE_OPTIONS):
                if opt_a in curr_fcfs_act or curr_fcfs_act in opt_a:
                    fcfs_act_idx = idx_a
                    break

            fcfs_action_rule = st.selectbox(
                "반영 방식 (사은품 출고 & 색인 방식)",
                ACTION_RULE_OPTIONS,
                index=fcfs_act_idx,
                key="ev_fcfs_act_rule",
                help="사은품을 별도 행으로 삽입할지, 행 추가 없이 하이라이트 색인만 처리할지 결정합니다."
            )
            fc_tc1, fc_tc2 = st.columns([1.5, 1])
            with fc_tc1:
                fcfs_write_col_on = st.checkbox(
                    "📝 발주서 특정 열에 [선착순] 텍스트 기재",
                    value=cur_fcfs.get("write_col_enabled", False),
                    key="ev_fcfs_write_col_on",
                    help="체크 시 지정한 엑셀 열(예: N열)에 [선착순] 문구가 기재됩니다."
                )
            with fc_tc2:
                fcfs_target_col = st.text_input(
                    "기재할 열 (영문 이니셜 or 열이름)",
                    value=cur_fcfs.get("write_target_col", "N"),
                    key="ev_fcfs_target_col",
                    disabled=not fcfs_write_col_on,
                    help="예: N (엑셀 N열에 입력), M, 비고 등"
                )

            fcfs_qty = st.number_input("지급 수량(개)", min_value=1, value=cur_fcfs.get("gift_quantity", 1), step=1, key="ev_fcfs_qty")

        fc_p1, fc_p2, fc_p3 = st.columns([1.8, 1.2, 1.4])
        with fc_p1:
            fcfs_pname = st.text_input("사은품 공식 상품명", value=cur_fcfs.get("gift_product_name", "[증정] 선착순 미니어처 15ml"), key="ev_fcfs_pname")
        with fc_p2:
            fcfs_opt = st.text_input("사은품 옵션명", value=cur_fcfs.get("gift_option_name", "증정용"), key="ev_fcfs_opt")
        with fc_p3:
            fcfs_scope_idx = 0 if cur_fcfs.get("highlight_scope", "both") == "both" else 1
            fcfs_hl_scope_text = st.radio(
                "색인 적용 범위",
                ["대상자들에 색인 (본품+사은품)", "이벤트에만 색인 (사은품만)"],
                index=fcfs_scope_idx,
                key="ev_fcfs_hl_scope"
            )
            fcfs_hl_scope = "both" if "대상자들에" in fcfs_hl_scope_text else "gift_only"

        fcfs_color, fcfs_no_color = render_google_style_color_selector(
            "선착순 사은품 색인(배경색)",
            cur_fcfs.get("color", "#FFF2CC"),
            key_prefix="ev_fcfs",
            default_no_color=cur_fcfs.get("no_color", False)
        )

        preview_bg = "#FFFFFF" if fcfs_no_color else fcfs_color
        preview_desc = "색인 없음 (흰색)" if fcfs_no_color else ("본품+사은품 모두 색인" if fcfs_hl_scope == "both" else "사은품 행만 색인 (본품은 흰색)")
        st.markdown(f"""
        <div style="padding: 7px 12px; background: {preview_bg}; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 13px; color: #2F3437; margin: 4px 0 10px 0;">
            <strong>[선착순 당첨 행 미리보기 ({preview_desc})]</strong> ORD-20261009 | 홍길동 | 주문 상품 (본품) / [증정] {fcfs_pname}
        </div>
        """, unsafe_allow_html=True)

        # 개별 카드 내부 즉시 5단계 이동 버튼 (사용자 요청 반영 ⭐)
        if st.button("👉 이 선착순 사은품 적용하고 5단계 이동", type="primary", key="btn_goto_step5_fcfs_direct"):
            fcfs_on = True
            top_apply_clicked = True

    # 2. 구매금액별 사은품 이벤트 (활성화 시 자동 펼침 ⭐)
    exp_amt = bool(cur_amt.get("enabled", False))
    with st.expander(f"💰 2. 구매금액별 사은품 이벤트 {'(✅ 활성)' if exp_amt else ''}", expanded=exp_amt):
        st.caption("주문 총 결제금액(본사 판매단가 기준)이 일정 금액 이상인 고객에게 사은품을 지급합니다.")
        amt_on = st.checkbox("구매금액별 사은품 적용하기", value=cur_amt.get("enabled", False), key="ev_amt_on")
        
        am_c1, am_c2 = st.columns([1.2, 1.2])
        with am_c1:
            amt_limit = st.number_input("기준 금액 (원 이상)", min_value=1000, value=cur_amt.get("min_amount", 50000), step=5000, key="ev_amt_limit")
            
            st.write("**합포장 및 대상 판정 기준**")
            saved_amt_targets = cur_amt.get("target_rule", ["주문번호"])
            if isinstance(saved_amt_targets, str):
                saved_amt_targets = [saved_amt_targets]

            achk_c1, achk_c2, achk_c3, achk_c4 = st.columns(4)
            with achk_c1:
                amt_k_oid = st.checkbox("주문번호", value=any("주문번호" in str(x) for x in saved_amt_targets), key="amt_chk_oid")
            with achk_c2:
                amt_k_name = st.checkbox("수취인명", value=any(k in str(x) for x in saved_amt_targets for k in ["수취인", "수령인", "구매자"]), key="amt_chk_name")
            with achk_c3:
                amt_k_phone = st.checkbox("연락처", value=any(k in str(x) for x in saved_amt_targets for k in ["연락처", "전화"]), key="amt_chk_phone")
            with achk_c4:
                amt_k_addr = st.checkbox("배송주소", value=any("주소" in str(x) for x in saved_amt_targets), key="amt_chk_addr")

            amt_crit_list = []
            if amt_k_oid: amt_crit_list.append("주문번호")
            if amt_k_name: amt_crit_list.append("수취인명")
            if amt_k_phone: amt_crit_list.append("연락처")
            if amt_k_addr: amt_crit_list.append("주소")
            if not amt_crit_list:
                amt_crit_list = ["주문번호"]

            amt_multiplier_on = st.checkbox("기준 금액 배수당 누적 지급 (예: 5만원당 1개, 10만원이면 2개)", value=any("배수" in str(x) for x in saved_amt_targets), key="amt_chk_multiplier")
            if amt_multiplier_on:
                amt_crit_list.append("배수당 누적")

            # 금액달성 고객 중 선착순 한정 지급 옵션 (사용자 요청 반영 ⭐)
            amt_fcfs_on = st.checkbox("기준 금액 달성 고객 중 선착순 한정 적용", value=cur_amt.get("fcfs_enabled", False), key="ev_amt_fcfs_on")
            amt_fcfs_limit = cur_amt.get("fcfs_limit", 30)
            if amt_fcfs_on:
                amt_fcfs_limit = st.number_input("선착순 인정 인원(명)", min_value=1, value=amt_fcfs_limit, step=1, key="ev_amt_fcfs_limit")
                st.caption(f"💡 결제금액이 {amt_limit:,}원 이상인 고객 중 주문시간이 빠른 선착순 **{amt_fcfs_limit}명**에게만 사은품 지급")

            st.caption(f"💡 합산 기준: **{' + '.join([c for c in amt_crit_list if c != '배수당 누적'])}** 동일 건 합산 금액으로 판정")

        with am_c2:
            curr_amt_act = cur_amt.get("gift_position", ACTION_RULE_OPTIONS[0])
            amt_act_idx = 0
            for idx_a, opt_a in enumerate(ACTION_RULE_OPTIONS):
                if opt_a in curr_amt_act or curr_amt_act in opt_a:
                    amt_act_idx = idx_a
                    break

            amt_action_rule = st.selectbox(
                "반영 방식 (금액달성 사은품 출고 & 색인 방식)",
                ACTION_RULE_OPTIONS,
                index=amt_act_idx,
                key="ev_amt_act_rule"
            )
            am_tc1, am_tc2 = st.columns([1.5, 1])
            with am_tc1:
                amt_write_col_on = st.checkbox(
                    "📝 발주서 특정 열에 [금액달성] 텍스트 기재",
                    value=cur_amt.get("write_col_enabled", False),
                    key="ev_amt_write_col_on",
                    help="체크 시 지정한 엑셀 열(예: N열)에 [금액달성] 문구가 기재됩니다."
                )
            with am_tc2:
                amt_target_col = st.text_input(
                    "기재할 열 (영문 이니셜 or 열이름) ",
                    value=cur_amt.get("write_target_col", "N"),
                    key="ev_amt_target_col",
                    disabled=not amt_write_col_on,
                    help="예: N (엑셀 N열에 입력), M, 비고 등"
                )

            amt_qty = st.number_input("기본 지급 수량(개)", min_value=1, value=cur_amt.get("gift_quantity", 1), step=1, key="ev_amt_qty")

        am_p1, am_p2, am_p3 = st.columns([1.8, 1.2, 1.4])
        with am_p1:
            amt_pname = st.text_input("사은품 공식 상품명 ", value=cur_amt.get("gift_product_name", "[증정] 금액달성 사은품"), key="ev_amt_pname")
        with am_p2:
            amt_opt = st.text_input("사은품 옵션명 ", value=cur_amt.get("gift_option_name", "1세트"), key="ev_amt_opt")
        with am_p3:
            amt_scope_idx = 0 if cur_amt.get("highlight_scope", "both") == "both" else 1
            amt_hl_scope_text = st.radio(
                "색인 적용 범위",
                ["대상자들에 색인 (본품+사은품)", "이벤트에만 색인 (사은품만)"],
                index=amt_scope_idx,
                key="ev_amt_hl_scope"
            )
            amt_hl_scope = "both" if "대상자들에" in amt_hl_scope_text else "gift_only"

        amt_color, amt_no_color = render_google_style_color_selector(
            "금액달성 사은품 색인(배경색)",
            cur_amt.get("color", "#FFF2CC"),
            key_prefix="ev_amt",
            default_no_color=cur_amt.get("no_color", False)
        )

        amt_preview_bg = "#FFFFFF" if amt_no_color else amt_color
        amt_preview_desc = "색인 없음 (흰색)" if amt_no_color else ("본품+사은품 모두 색인" if amt_hl_scope == "both" else "사은품 행만 색인 (본품은 흰색)")
        st.markdown(f"""
        <div style="padding: 7px 12px; background: {amt_preview_bg}; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 13px; color: #2F3437; margin: 4px 0 10px 0;">
            <strong>[금액달성 당첨 행 미리보기 ({amt_preview_desc})]</strong> ORD-20261009 | 홍길동 | 주문 상품 (본품) / [증정] {amt_pname}
        </div>
        """, unsafe_allow_html=True)

        if st.button("👉 이 금액달성 사은품 적용하고 5단계 이동", type="primary", key="btn_goto_step5_amt_direct"):
            amt_on = True
            top_apply_clicked = True

    # 3. 특정 품목 구매 사은품 (활성화 시 자동 펼침 ⭐)
    exp_spec = bool(cur_spec.get("enabled", False))
    with st.expander(f"🎁 3. 특정 품목 구매 사은품 이벤트 {'(✅ 활성)' if exp_spec else ''}", expanded=exp_spec):
        st.caption("특정 상품 또는 옵션을 구매한 고객에게 전용 사은품을 지급합니다.")
        spec_on = st.checkbox("특정 품목 사은품 적용하기", value=cur_spec.get("enabled", False), key="ev_spec_on")
        
        sp_c1, sp_c2 = st.columns([1.2, 1.2])
        with sp_c1:
            spec_kw = st.text_input("조건 품목 검색 키워드 (예: 세트, 참치)", value=cur_spec.get("keyword", ""), key="ev_spec_kw")
            
            st.write("**지급 대상 판정 기준**")
            saved_spec_targets = cur_spec.get("target_rule", ["주문번호"])
            if isinstance(saved_spec_targets, str):
                saved_spec_targets = [saved_spec_targets]

            schk_c1, schk_c2, schk_c3, schk_c4 = st.columns(4)
            with schk_c1:
                spec_k_oid = st.checkbox("주문번호", value=any("주문번호" in str(x) for x in saved_spec_targets), key="spec_chk_oid")
            with schk_c2:
                spec_k_name = st.checkbox("수취인명", value=any(k in str(x) for x in saved_spec_targets for k in ["수취인", "수령인", "구매자"]), key="spec_chk_name")
            with schk_c3:
                spec_k_phone = st.checkbox("연락처", value=any(k in str(x) for x in saved_spec_targets for k in ["연락처", "전화"]), key="spec_chk_phone")
            with schk_c4:
                spec_k_addr = st.checkbox("배송주소", value=any("주소" in str(x) for x in saved_spec_targets), key="spec_chk_addr")

            spec_crit_list = []
            if spec_k_oid: spec_crit_list.append("주문번호")
            if spec_k_name: spec_crit_list.append("수취인명")
            if spec_k_phone: spec_crit_list.append("연락처")
            if spec_k_addr: spec_crit_list.append("주소")
            if not spec_crit_list:
                spec_crit_list = ["주문번호"]

            spec_ratio_on = st.checkbox("해당 품목 구매 수량과 1:1 비례 지급 (구매수량만큼 사은품 증정)", value=any("1:1" in str(x) for x in saved_spec_targets), key="spec_chk_ratio")
            if spec_ratio_on:
                spec_crit_list.append("구매 수량과 1:1")

            st.caption(f"💡 지급 기준: **{' + '.join([c for c in spec_crit_list if c != '구매 수량과 1:1'])}** 기준 판정")

        with sp_c2:
            curr_spec_act = cur_spec.get("gift_position", ACTION_RULE_OPTIONS[0])
            spec_act_idx = 0
            for idx_a, opt_a in enumerate(ACTION_RULE_OPTIONS):
                if opt_a in curr_spec_act or curr_spec_act in opt_a:
                    spec_act_idx = idx_a
                    break

            spec_action_rule = st.selectbox(
                "반영 방식 (특정품목 사은품 출고 & 색인 방식)",
                ACTION_RULE_OPTIONS,
                index=spec_act_idx,
                key="ev_spec_act_rule"
            )
            sp_tc1, sp_tc2 = st.columns([1.5, 1])
            with sp_tc1:
                spec_write_col_on = st.checkbox(
                    "📝 발주서 특정 열에 [품목증정] 텍스트 기재",
                    value=cur_spec.get("write_col_enabled", False),
                    key="ev_spec_write_col_on",
                    help="체크 시 지정한 엑셀 열(예: N열)에 [품목증정] 문구가 기재됩니다."
                )
            with sp_tc2:
                spec_target_col = st.text_input(
                    "기재할 열 (영문 이니셜 or 열이름)  ",
                    value=cur_spec.get("write_target_col", "N"),
                    key="ev_spec_target_col",
                    disabled=not spec_write_col_on,
                    help="예: N (엑셀 N열에 입력), M, 비고 등"
                )

            spec_qty = st.number_input("지급 수량(개) ", min_value=1, value=cur_spec.get("gift_quantity", 1), step=1, key="ev_spec_qty")

        sp_p1, sp_p2, sp_p3 = st.columns([1.8, 1.2, 1.4])
        with sp_p1:
            spec_pname = st.text_input("사은품 공식 상품명  ", value=cur_spec.get("gift_product_name", "[사은품] 전용 세탁망"), key="ev_spec_pname")
        with sp_p2:
            spec_opt = st.text_input("사은품 옵션명  ", value=cur_spec.get("gift_option_name", "단품"), key="ev_spec_opt")
        with sp_p3:
            spec_scope_idx = 0 if cur_spec.get("highlight_scope", "both") == "both" else 1
            spec_hl_scope_text = st.radio(
                "색인 적용 범위  ",
                ["대상자들에 색인 (본품+사은품)", "이벤트에만 색인 (사은품만)"],
                index=spec_scope_idx,
                key="ev_spec_hl_scope"
            )
            spec_hl_scope = "both" if "대상자들에" in spec_hl_scope_text else "gift_only"

        spec_color, spec_no_color = render_google_style_color_selector(
            "특정품목 사은품 색인(배경색)",
            cur_spec.get("color", "#E2EFDA"),
            key_prefix="ev_spec",
            default_no_color=cur_spec.get("no_color", False)
        )

        spec_preview_bg = "#FFFFFF" if spec_no_color else spec_color
        spec_preview_desc = "색인 없음 (흰색)" if spec_no_color else ("본품+사은품 모두 색인" if spec_hl_scope == "both" else "사은품 행만 색인 (본품은 흰색)")
        st.markdown(f"""
        <div style="padding: 7px 12px; background: {spec_preview_bg}; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 13px; color: #2F3437; margin: 4px 0 10px 0;">
            <strong>[특정품목 증정 행 미리보기 ({spec_preview_desc})]</strong> ORD-20261009 | 홍길동 | 주문 상품 (본품) / [사은품] {spec_pname}
        </div>
        """, unsafe_allow_html=True)

        if st.button("👉 이 특정품목 사은품 적용하고 5단계 이동", type="primary", key="btn_goto_step5_spec_direct"):
            spec_on = True
            top_apply_clicked = True

    # 4. 자유 추가 커스텀 사은품 (추가 프로모션 행사 ⭐)
    if "custom_events_list" not in st.session_state:
        st.session_state["custom_events_list"] = [dict(x) for x in cfg_cur.get("custom_events", [])]

    custom_events_cur_list = st.session_state["custom_events_list"]
    with st.expander(f"✨ 4. 자유 추가 커스텀 사은품 (추가 프로모션 {len(custom_events_cur_list)}건)", expanded=bool(custom_events_cur_list)):
        st.caption("기본 3가지 프로모션 외에 전 고객 증정, 추가 선착순, 특정 금액/키워드 증정 등 새로운 사은품 이벤트를 자유롭게 추가하고 관리하세요.")
        
        c_add_btn = st.button("➕ 새 사은품 이벤트 추가", type="secondary", key="btn_add_custom_event")
        if c_add_btn:
            new_idx = len(custom_events_cur_list) + 1
            custom_events_cur_list.append({
                "title": f"사은품 이벤트 {new_idx}",
                "enabled": True,
                "condition_type": "전체 주문 증정",
                "limit": 10,
                "min_amount": 30000,
                "keyword": "",
                "gift_product_name": "[증정] 특별 사은품",
                "gift_option_name": "단품",
                "gift_quantity": 1,
                "target_rule": ["수취인명", "연락처"],
                "gift_position": ACTION_RULE_OPTIONS[0],
                "tag_display_rule": "텍스트 표기 없음 (화면 및 엑셀 배경색 하이라이트 색인만 처리 ⭐)"
            })
            st.rerun()

        custom_events_data = []
        del_custom_idx = None

        for c_i, c_item in enumerate(custom_events_cur_list):
            st.markdown(f"**📌 사은품 이벤트 #{c_i+1}: {c_item.get('title', '이벤트')}**")
            c_box = st.container(border=True)
            with c_box:
                c_top1, c_top2, c_top3 = st.columns([2, 1, 1])
                with c_top1:
                    c_title = st.text_input("이벤트 명칭", value=c_item.get("title", f"사은품 이벤트 {c_i+1}"), key=f"c_title_{c_i}")
                with c_top2:
                    c_on = st.checkbox("적용 활성화", value=c_item.get("enabled", True), key=f"c_on_{c_i}")
                with c_top3:
                    if st.button("🗑️ 삭제", key=f"c_del_{c_i}", type="secondary"):
                        del_custom_idx = c_i

                c_cond_types = ["전체 주문 증정", "선착순 주문 증정", "구매금액 달성 증정", "특정 품목/키워드 증정"]
                curr_c_cond = c_item.get("condition_type", "전체 주문 증정")
                c_cond_idx = 0
                for c_k_idx, c_cond_name in enumerate(c_cond_types):
                    if c_cond_name in curr_c_cond or curr_c_cond in c_cond_name:
                        c_cond_idx = c_k_idx
                        break
                
                c_col_a, c_col_b = st.columns(2)
                with c_col_a:
                    c_cond = st.selectbox("증정 조건 유형", c_cond_types, index=c_cond_idx, key=f"c_cond_{c_i}")
                    c_limit = c_item.get("limit", 10)
                    c_min_amt = c_item.get("min_amount", 30000)
                    c_kw = c_item.get("keyword", "")

                    if "선착순" in c_cond:
                        c_limit = st.number_input("선착순 인정 인원(명)", min_value=1, value=c_limit, step=1, key=f"c_limit_{c_i}")
                    elif "금액" in c_cond:
                        c_min_amt = st.number_input("기준 결제금액 (원 이상)", min_value=1000, value=c_min_amt, step=5000, key=f"c_min_amt_{c_i}")
                    elif "특정 품목" in c_cond or "키워드" in c_cond:
                        c_kw = st.text_input("조건 품목 키워드 (예: 베개, 세트)", value=c_kw, key=f"c_kw_{c_i}")

                    # 판정 기준 체크박스
                    st.write("**동일 고객 판정 기준**")
                    c_saved_targets = c_item.get("target_rule", ["수취인명", "연락처"])
                    if isinstance(c_saved_targets, str):
                        c_saved_targets = [c_saved_targets]
                    
                    c_chk1, c_chk2, c_chk3, c_chk4 = st.columns(4)
                    with c_chk1:
                        c_k_oid = st.checkbox("주문번호", value=any("주문번호" in str(x) for x in c_saved_targets), key=f"c_k_oid_{c_i}")
                    with c_chk2:
                        c_k_name = st.checkbox("수취인명", value=any(k in str(x) for x in c_saved_targets for k in ["수취인", "수령인", "구매자"]), key=f"c_k_name_{c_i}")
                    with c_chk3:
                        c_k_phone = st.checkbox("연락처", value=any(k in str(x) for x in c_saved_targets for k in ["연락처", "전화"]), key=f"c_k_phone_{c_i}")
                    with c_chk4:
                        c_k_addr = st.checkbox("배송주소", value=any("주소" in str(x) for x in c_saved_targets), key=f"c_k_addr_{c_i}")

                    c_target_rule = []
                    if c_k_oid: c_target_rule.append("주문번호")
                    if c_k_name: c_target_rule.append("수취인명")
                    if c_k_phone: c_target_rule.append("연락처")
                    if c_k_addr: c_target_rule.append("주소")
                    if not c_target_rule:
                        c_target_rule = ["수취인명", "연락처"]

                with c_col_b:
                    curr_c_act = c_item.get("gift_position", ACTION_RULE_OPTIONS[0])
                    c_act_idx = 0
                    for idx_a, opt_a in enumerate(ACTION_RULE_OPTIONS):
                        if opt_a in curr_c_act or curr_c_act in opt_a:
                            c_act_idx = idx_a
                            break
                    c_action_rule = st.selectbox("반영 방식", ACTION_RULE_OPTIONS, index=c_act_idx, key=f"c_act_rule_{c_i}")
                    c_tc1, c_tc2 = st.columns([1.5, 1])
                    with c_tc1:
                        c_write_col_on = st.checkbox(
                            f"📝 발주서 특정 열에 [{c_title}] 텍스트 기재",
                            value=c_item.get("write_col_enabled", False),
                            key=f"c_write_col_on_{c_i}",
                            help="체크 시 지정한 엑셀 열에 문구가 기재됩니다."
                        )
                    with c_tc2:
                        c_target_col = st.text_input(
                            "기재할 열 (영문 이니셜 or 열이름)",
                            value=c_item.get("write_target_col", "N"),
                            key=f"c_target_col_{c_i}",
                            disabled=not c_write_col_on,
                            help="예: N (엑셀 N열에 입력), M, 비고 등"
                        )

                    c_qty = st.number_input("사은품 지급 수량(개)", min_value=1, value=c_item.get("gift_quantity", 1), step=1, key=f"c_qty_{c_i}")

                c_g_c1, c_g_c2, c_g_c3 = st.columns([1.8, 1.2, 1.4])
                with c_g_c1:
                    c_pname = st.text_input("사은품 공식 상품명", value=c_item.get("gift_product_name", "[증정] 추가 사은품"), key=f"c_pname_{c_i}")
                with c_g_c2:
                    c_opt = st.text_input("사은품 옵션명", value=c_item.get("gift_option_name", "단품"), key=f"c_opt_{c_i}")
                with c_g_c3:
                    c_scope_idx = 0 if c_item.get("highlight_scope", "both") == "both" else 1
                    c_hl_scope_text = st.radio(
                        "색인 적용 범위",
                        ["대상자들에 색인 (본품+사은품)", "이벤트에만 색인 (사은품만)"],
                        index=c_scope_idx,
                        key=f"c_hl_scope_{c_i}"
                    )
                    c_hl_scope = "both" if "대상자들에" in c_hl_scope_text else "gift_only"

                c_color, c_no_color = render_google_style_color_selector(
                    f"{c_title} 색인(배경색)",
                    c_item.get("color", "#FFF2CC"),
                    key_prefix=f"c_{c_i}",
                    default_no_color=c_item.get("no_color", False)
                )

                c_preview_bg = "#FFFFFF" if c_no_color else c_color
                c_preview_desc = "색인 없음 (흰색)" if c_no_color else ("본품+사은품 모두 색인" if c_hl_scope == "both" else "사은품 행만 색인 (본품은 흰색)")
                st.markdown(f"""
                <div style="padding: 7px 12px; background: {c_preview_bg}; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 13px; color: #2F3437; margin: 4px 0 10px 0;">
                    <strong>[{c_title} 당첨 행 미리보기 ({c_preview_desc})]</strong> ORD-20261009 | 홍길동 | 주문 상품 (본품) / [증정] {c_pname}
                </div>
                """, unsafe_allow_html=True)

                c_act_type, c_pos, c_disp = parse_action_rule(c_action_rule)
                custom_events_data.append({
                    "title": c_title,
                    "enabled": c_on,
                    "condition_type": c_cond,
                    "limit": c_limit,
                    "min_amount": c_min_amt,
                    "keyword": c_kw,
                    "gift_product_name": c_pname,
                    "gift_option_name": c_opt,
                    "gift_quantity": c_qty,
                    "target_rule": c_target_rule,
                    "action_type": c_act_type,
                    "gift_position": c_pos,
                    "tag_display_rule": c_disp,
                    "write_col_enabled": c_write_col_on,
                    "write_target_col": c_target_col.strip() if c_target_col else "N",
                    "color": c_color,
                    "no_color": c_no_color,
                    "highlight_scope": c_hl_scope
                })

        if del_custom_idx is not None:
            custom_events_cur_list.pop(del_custom_idx)
            st.session_state["custom_events_list"] = custom_events_cur_list
            st.rerun()

    st.markdown("---")
    # 하단 메인 액션 바: 템플릿 저장 (에메랄드 그린) & 5단계 최종 발주서 생성 (프라이머리 네이비)
    bottom_save_template_clicked = False
    bottom_apply_clicked = False

    s4_bot1, s4_bot2 = st.columns([1.4, 2])
    with s4_bot1:
        if st.button("💾 현재 이벤트 설정 템플릿 저장하기", type="primary", use_container_width=True, key="btn_save_event_config_bottom"):
            bottom_save_template_clicked = True
    with s4_bot2:
        if st.button("👉 5단계 최종 발주서 & 정산표 생성하기", type="primary", use_container_width=True, key="btn_goto_step5_from4"):
            bottom_apply_clicked = True

    save_template_triggered = (top_save_template_clicked or bottom_save_template_clicked)
    apply_to_step5_triggered = (top_apply_clicked or bottom_apply_clicked)

    # 전체 저장 로직 및 실시간 세션 보존 ⭐
    fcfs_act_type, fcfs_pos, fcfs_tag_rule = parse_action_rule(fcfs_action_rule)
    amt_act_type, amt_pos, amt_tag_rule = parse_action_rule(amt_action_rule)
    spec_act_type, spec_pos, spec_tag_rule = parse_action_rule(spec_action_rule)

    new_event_config = {
        "color_event": fcfs_color if 'fcfs_color' in locals() else "#FFF2CC",
        "color_gift": spec_color if 'spec_color' in locals() else "#E2EFDA",
        "first_come_first_served": {
            "enabled": fcfs_on,
            "limit": fcfs_limit,
            "gift_product_name": fcfs_pname,
            "gift_option_name": fcfs_opt,
            "gift_quantity": fcfs_qty,
            "target_rule": fcfs_crit_list,
            "action_type": fcfs_act_type,
            "tag_display_rule": fcfs_tag_rule,
            "gift_position": fcfs_pos,
            "write_col_enabled": fcfs_write_col_on if 'fcfs_write_col_on' in locals() else False,
            "write_target_col": fcfs_target_col.strip() if 'fcfs_target_col' in locals() and fcfs_target_col else "N",
            "color": fcfs_color,
            "no_color": fcfs_no_color if 'fcfs_no_color' in locals() else False,
            "highlight_scope": fcfs_hl_scope if 'fcfs_hl_scope' in locals() else "both"
        },
        "min_amount": {
            "enabled": amt_on,
            "min_amount": amt_limit,
            "fcfs_enabled": amt_fcfs_on,
            "fcfs_limit": amt_fcfs_limit,
            "gift_product_name": amt_pname,
            "gift_option_name": amt_opt,
            "gift_quantity": amt_qty,
            "target_rule": amt_crit_list,
            "action_type": amt_act_type,
            "tag_display_rule": amt_tag_rule,
            "gift_position": amt_pos,
            "write_col_enabled": amt_write_col_on if 'amt_write_col_on' in locals() else False,
            "write_target_col": amt_target_col.strip() if 'amt_target_col' in locals() and amt_target_col else "N",
            "color": amt_color,
            "no_color": amt_no_color if 'amt_no_color' in locals() else False,
            "highlight_scope": amt_hl_scope if 'amt_hl_scope' in locals() else "both"
        },
        "special_product": {
            "enabled": spec_on,
            "keyword": spec_kw,
            "gift_product_name": spec_pname,
            "gift_option_name": spec_opt,
            "gift_quantity": spec_qty,
            "target_rule": spec_crit_list,
            "action_type": spec_act_type,
            "tag_display_rule": spec_tag_rule,
            "gift_position": spec_pos,
            "write_col_enabled": spec_write_col_on if 'spec_write_col_on' in locals() else False,
            "write_target_col": spec_target_col.strip() if 'spec_target_col' in locals() and spec_target_col else "N",
            "color": spec_color,
            "no_color": spec_no_color if 'spec_no_color' in locals() else False,
            "highlight_scope": spec_hl_scope if 'spec_hl_scope' in locals() else "both"
        },
        "custom_events": custom_events_data if 'custom_events_data' in locals() else []
    }
    st.session_state["event_config"] = new_event_config

    if save_template_triggered:
        tmpl_to_save = input_ev_tmpl_title.strip() if 'input_ev_tmpl_title' in locals() and input_ev_tmpl_title.strip() else f"{active_sel_name}_{active_b_name}_이벤트"
        save_event_template(tmpl_to_save, new_event_config)
        
        # 셀러 매핑 템플릿에도 연동 저장
        active_seller_tmpl = st.session_state.get("active_seller_tmpl_key")
        if active_seller_tmpl:
            s_data = load_seller_mappings()
            if active_seller_tmpl in s_data:
                s_data[active_seller_tmpl]["event_config"] = new_event_config
                save_seller_mappings(s_data)
                
        st.session_state["pending_event_tmpl_choice"] = tmpl_to_save
        st.session_state["last_ev_sync_context"] = None
        st.success(f"✨ '{tmpl_to_save}' 이벤트 템플릿이 성공적으로 저장되었습니다! 상단 목록에서 즉시 확인하실 수 있습니다.")
        st.rerun()
    elif apply_to_step5_triggered:
        st.session_state["event_config_confirmed"] = True
        st.session_state["event_skip_mode"] = False
        st.session_state["active_step_idx"] = 4
        st.rerun()

    # 이벤트 규칙 JSON 프리셋 백업/불러오기
    with st.expander("📁 이벤트 규칙 프리셋 백업 및 불러오기", expanded=False):
        ev_down_col, ev_up_col = st.columns(2)
        with ev_down_col:
            st.download_button(
                label="📥 현재 이벤트 설정 다운로드 (.json)",
                data=json.dumps(st.session_state.get("event_config", default_event_config), ensure_ascii=False, indent=2),
                file_name=f"event_rules_preset_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json",
                use_container_width=True
            )
        with ev_up_col:
            up_ev_json = st.file_uploader("📤 이벤트 설정 JSON 불러오기", type=["json", "JSON"], key="uploader_event_preset")
            if up_ev_json is not None:
                if st.button("🔄 이벤트 프리셋 적용", type="secondary", key="btn_apply_ev_preset"):
                    try:
                        loaded_ev = json.loads(up_ev_json.read().decode("utf-8"))
                        st.session_state["event_config"] = loaded_ev
                        st.session_state["event_config_confirmed"] = True
                        st.success("이벤트 프리셋이 적용되었습니다!")
                        st.rerun()
                    except Exception as e_ev:
                        st.error(f"오류: {e_ev}")

# =========================================================================
# STEP 5: 최종 발주서 변환 & 다운로드
# =========================================================================
elif cur_step_idx == 4:
    st.markdown("""
    <div class="notion-step-banner">5️⃣ 5단계. 최종 발주서 & 정산표 다운로드</div>
    <div class="notion-step-desc-big">모든 매핑, 본사 판매단가, 프로모션 색인/사은품이 완벽히 적용된 브랜드사 전달용 발주서와 일자별 품목/배송 정산표를 확인하고 다운로드하세요.</div>
    """, unsafe_allow_html=True)

    if "current_seller_df" in st.session_state and "user_col_map" in st.session_state:
        df_seller = st.session_state.current_seller_df
        seller_file_name = st.session_state.current_seller_file
        user_col_map = st.session_state.user_col_map
        
        active_seller_tmpl = st.session_state.get("active_seller_tmpl_key", seller_file_name)
        seller_map_data = get_seller_mapping(active_seller_tmpl) or get_seller_mapping(seller_file_name) or {}
        pm_dict = st.session_state.get("active_product_mapping", seller_map_data.get("product_mappings", {}))

        active_brand_name = st.session_state.get("active_brand", "호무로")
        active_brand_cols = st.session_state.get("active_brand_cols")
        if not active_brand_cols:
            b_tmpl_data = brand_templates.get(active_brand_name, {})
            active_brand_cols = b_tmpl_data.get("columns", [])
        if not active_brand_cols and brand_templates:
            first_b_k = list(brand_templates.keys())[0]
            active_brand_cols = brand_templates.get(first_b_k, {}).get("columns", [])
        active_brand_template = {"columns": active_brand_cols if active_brand_cols else []}

        # 커스텀 일자 및 커스텀 판매처 확인
        active_custom_order_date = st.session_state.get("active_custom_order_date")
        if not active_custom_order_date and seller_map_data.get("custom_order_date"):
            active_custom_order_date = seller_map_data.get("custom_order_date")

        active_force_custom_date = st.session_state.get("active_force_custom_date")
        if active_force_custom_date is None:
            active_force_custom_date = seller_map_data.get("force_custom_date", True)

        active_custom_seller = st.session_state.get("active_custom_seller_name")
        if not active_custom_seller and seller_map_data.get("custom_seller_name"):
            active_custom_seller = seller_map_data.get("custom_seller_name")

        active_seller_name = active_custom_seller if active_custom_seller else seller_map_data.get("seller_name", extract_clean_seller_from_filename(seller_file_name) or seller_file_name)

        # 1. 정규화 및 단가 기반 금액 계산
        norm_df = process_seller_dataframe(
            df_seller, 
            seller_file_name, 
            user_col_map, 
            pm_dict,
            custom_order_date=active_custom_order_date,
            force_custom_date=active_force_custom_date,
            custom_seller_name=active_custom_seller
        )
        
        # 활성화된 이벤트 설정 적용
        active_event_cfg = st.session_state.get("event_config", default_event_config)
        color_event = active_event_cfg.get("color_event", "#FFF2CC")
        color_gift = active_event_cfg.get("color_gift", "#E2EFDA")
        processed_df = apply_event_rules(norm_df, active_event_cfg)

        # 1-1. 일자별 발주 데이터 누적 관리 (사용자 요청: 동일 템플릿 다음날 누적 지원 ⭐)
        accum_df = load_accumulated_orders(active_seller_name, active_brand_name)
        has_history = not accum_df.empty
        current_dates = sorted(list(processed_df["order_date"].astype(str).unique()))
        history_file_path = os.path.abspath(get_order_history_file(active_seller_name, active_brand_name))

        st.markdown(f"""
        <div style="background-color: #F8F9FA; border: 1px solid #E2E2DF; border-radius: 8px; padding: 12px 18px; margin-bottom: 12px;">
            <div style="font-size: 15px; font-weight: 600; color: #1E3D59; margin-bottom: 4px;">
                📅 템플릿 연동 일자별 누적 관리: <span style="color: #2E5B88;">{active_seller_name} + {active_brand_name}</span>
            </div>
            <div style="font-size: 13px; color: #55534E; line-height: 1.5;">
                오늘 발주 데이터(<strong>{', '.join(current_dates)}</strong>)를 누적 보관소에 저장해 두면, 다음날 새로운 발주서를 불러올 때 일자별 수량표 옆에 날짜가 자동으로 나란히 추가되어 누적 집계됩니다.<br>
                <span style="font-size: 12px; color: #888;">📁 물리적 저장 위치: <code style="color: #2E5B88;">{history_file_path}</code></span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.session_state.get("accum_saved_banner"):
            st.success(f"✅ {st.session_state['accum_saved_banner']}")
            st.info(f"💡 **동작 원리 안내**: 현재 발주서의 날짜({', '.join(current_dates)})가 보관소에 안전하게 기록되었습니다. 지금은 오늘 날짜 데이터만 조회 중이므로 수량표 열이 1개지만, **내일 새 발주서를 업로드하시면 수량표에 [{', '.join(current_dates)}, 내일날짜]가 가로 열로 나란히 누적 집계**됩니다.")

        acc_col1, acc_col2 = st.columns([3, 1.5])
        with acc_col1:
            use_accum = st.checkbox(
                "📅 이전 일자 발주 데이터 누적하여 함께 보기 (수량표에 모든 일자 가로 표시)",
                value=has_history,
                help="체크 시 이전에 처리했던 날짜의 발주 데이터가 함께 병합되어, 수량표의 날짜 열(Columns)에 가로로 나란히 표시됩니다."
            )
            if has_history:
                hist_dates = sorted(list(accum_df["order_date"].astype(str).unique()))
                st.caption(f"📁 현재 보관된 누적 일자: **{', '.join(hist_dates)}** (총 {len(accum_df)}행, 주문 {accum_df['order_id'].nunique()}건)")
                with st.expander("📂 보관소에 저장된 누적 발주 내역 열어보기"):
                    acc_view = accum_df[["order_date", "seller_name", "order_id", "receiver_name", "product_name", "option_name", "quantity", "price"]].head(30).copy()
                    acc_view.index = range(1, len(acc_view) + 1)
                    st.dataframe(acc_view, use_container_width=True)
            else:
                st.caption("ℹ️ 아직 보관된 이전 발주 내역이 없습니다. [💾 현재 발주 누적 저장]을 누르면 다음 발주부터 자동 누적됩니다.")

        with acc_col2:
            h_b1, h_b2 = st.columns(2)
            with h_b1:
                if st.button("💾 현재 발주 누적 저장", type="secondary", use_container_width=True, help="현재 발주 데이터를 누적 보관소에 저장합니다."):
                    save_accumulated_orders(active_seller_name, active_brand_name, processed_df)
                    st.session_state["accum_saved_banner"] = f"{', '.join(current_dates)} 발주 데이터 {len(processed_df)}행(주문 {processed_df['order_id'].nunique()}건)이 누적 보관소에 성공적으로 저장되었습니다!"
                    st.rerun()
            with h_b2:
                if has_history:
                    if st.button("🗑️ 누적 초기화", use_container_width=True, help="보관된 이전 누적 데이터를 모두 삭제하고 현재 파일만 봅니다."):
                        clear_accumulated_orders(active_seller_name, active_brand_name)
                        st.session_state.pop("accum_saved_banner", None)
                        st.info("누적 데이터가 초기화되었습니다.")
                        st.rerun()

        # 누적 사용 여부에 따른 집계 대상 데이터프레임 결정
        if use_accum and has_history:
            combined_df = pd.concat([accum_df, processed_df], ignore_index=True)
            subset_cols = [c for c in ["order_id", "product_name", "option_name", "order_date"] if c in combined_df.columns]
            combined_df.drop_duplicates(subset=subset_cols, keep="last", inplace=True)
            eval_df = combined_df
        else:
            eval_df = processed_df

        # 3단계 품목 정렬 순서 가져오기 (5단계 일자별 수량표에 100% 동일 순서 반영 ⭐)
        active_product_order = st.session_state.get("active_product_order")
        if not active_product_order and seller_map_data:
            active_product_order = seller_map_data.get("product_order", [])

        custom_order_map = {}
        if active_product_order:
            for s_idx, k_p in enumerate(active_product_order):
                info = pm_dict.get(k_p, {})
                bp = str(info.get("brand_product", "")).strip() or k_p.split("||")[0].strip()
                bo = str(info.get("brand_option", "")).strip()
                if (bp, bo) not in custom_order_map:
                    custom_order_map[(bp, bo)] = s_idx + 1
                if bp not in custom_order_map:
                    custom_order_map[bp] = s_idx + 1
        elif pm_dict:
            for k_p, info in pm_dict.items():
                bp = str(info.get("brand_product", "")).strip() or k_p.split("||")[0].strip()
                bo = str(info.get("brand_option", "")).strip()
                s_ord = info.get("sort_order")
                if s_ord is not None:
                    if (bp, bo) not in custom_order_map:
                        custom_order_map[(bp, bo)] = int(s_ord)
                    if bp not in custom_order_map:
                        custom_order_map[bp] = int(s_ord)

        # 2. 결과 시트 생성 (일체형 수량표 + 배송정산표 + 마스터 DB)
        brand_df = build_brand_order_sheet(processed_df, active_brand_template, active_event_cfg)
        daily_qty_df, daily_shipping_df, enriched_calc_df = build_daily_summary_sheets(
            eval_df,
            shipping_config,
            event_config=active_event_cfg,
            custom_item_order=custom_order_map
        )

        # 마스터 Raw DB: 배송비 구분 및 합포장 총금액 정렬
        master_raw_df = enriched_calc_df.drop(columns=["order_time_obj", "gifts_to_add", "bundle_id"], errors="ignore").copy()
        if "event_tags" in master_raw_df.columns:
            master_raw_df["event_tags"] = master_raw_df["event_tags"].apply(lambda tags: ", ".join(tags) if isinstance(tags, (list, tuple, set)) and tags else (str(tags) if tags else "해당없음"))
        if "is_jeju" in master_raw_df.columns:
            master_raw_df["is_jeju"] = master_raw_df["is_jeju"].apply(lambda j: "제주/도서산간" if j else "일반지역")

        front_cols = ["order_date", "seller_name", "order_id", "receiver_name", "receiver_phone", "full_address", "product_name", "option_name", "unit_price", "quantity", "price", "배송비구분", "묶음주문합계", "부과배송비", "event_tags", "shipping_memo"]
        valid_front = [c for c in front_cols if c in master_raw_df.columns]
        other_cols = [c for c in master_raw_df.columns if c not in valid_front]
        master_raw_df = master_raw_df[valid_front + other_cols]

        # 3. 요약 지표 카드
        m1, m2, m3, m4, m5 = st.columns(5)
        total_orders = eval_df["order_id"].nunique()
        total_items = eval_df["quantity"].sum()
        jeju_count = eval_df[eval_df["is_jeju"]]["order_id"].nunique()
        event_winner_count = sum(1 for tags in eval_df["event_tags"] if any("선착순" in t for t in tags))
        total_revenue = eval_df["price"].sum()

        m1.metric("총 주문 건수", f"{total_orders:,} 건")
        m2.metric("총 출고 수량", f"{total_items:,} 개")
        m3.metric("제주/도서 건수", f"{jeju_count:,} 건")
        m4.metric("선착순 대상", f"{event_winner_count:,} 건")
        m5.metric("총 정산 매출", f"{total_revenue:,} 원")

        # 4. 엑셀 다운로드 (커스텀 색상 및 통합 수량표 반영)
        st.write(" ")
        excel_bytes = export_styled_excel(brand_df, daily_qty_df, daily_shipping_df, master_raw_df, color_event=color_event, color_gift=color_gift)
        now_str = datetime.now().strftime("%Y%m%d_%H%M")

        dl_col1, dl_col2 = st.columns([2.5, 1])
        with dl_col1:
            st.markdown('<div class="final-download-btn-pulse">', unsafe_allow_html=True)
            st.download_button(
                label=f"📥 [{active_brand_name}] 최종 발주서 & 일자별 품목수량표 다운로드 (.xlsx)",
                data=excel_bytes,
                file_name=f"{active_brand_name}_최종발주서_{now_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="btn_download_final_order_sheet"
            )
            st.markdown('</div>', unsafe_allow_html=True)
        with dl_col2:
            st.caption(f"✅ 브랜드사('{active_brand_name}') 발주서, 하단 배송정산 통합 수량표, 및 마스터 DB가 포함되어 있습니다.")

        # 5. 상세 결과 탭 (사용자 요청: 수량표 일체형 통합 ⭐)
        res_t1, res_t2, res_t3 = st.tabs([
            "📊 1. 일자별 품목 수량 통합표 (날짜 열 / 품목 행)",
            f"🏢 2. {active_brand_name} 제출 최종 발주서",
            "💾 3. 셀러 마스터 Raw DB (배송비 구분 & 유료배송 조회)"
        ])

        with res_t1:
            st.subheader("📊 일자별 품목 수량 통합 집계표")
            st.caption("각 품목의 **판매단가, 일자별 출고 수량, 총 수량, 총 매출액(원)**과 하단의 **주문건수 / 유료배송 / 무료배송 / 제주 / 이벤트** 요약 행이 가로 날짜 열에 맞춰 일체형으로 배치된 정산표입니다. (💡 3단계에서 지정한 순서 및 품목 번호 순서 그대로 행이 정렬되어 출력됩니다.)")

            def highlight_qty_rows(row):
                pname = str(row["브랜드상품명"])
                if pname == "[전체 합계]":
                    return ["background-color: #EBF3FA; font-weight: bold; color: #1E3D59;"] * len(row)
                elif pname in ["주문건수", "유료배송", "무료배송", "제주"]:
                    return ["background-color: #F8F9FA; font-weight: bold; color: #2F3437;"] * len(row)
                elif pname == "이벤트" or (pname.startswith("[") and pname != "[전체 합계]"):
                    return ["background-color: #FFF9E6; font-weight: bold; color: #B78103;"] * len(row)
                elif pname == "":
                    return ["background-color: #FFFFFF;"] * len(row)
                return [""] * len(row)

            disp_qty_df = daily_qty_df.copy()
            for col in disp_qty_df.columns:
                if col in ["브랜드상품명", "옵션명"]:
                    disp_qty_df[col] = disp_qty_df[col].fillna("").astype(str)
                elif col == "판매단가":
                    disp_qty_df[col] = disp_qty_df[col].apply(lambda v: "-" if (pd.isna(v) or v in ["", 0, None, "-"]) else (f"{int(v):,}원" if isinstance(v, (int, float)) else str(v)))
                elif col == "총매출액":
                    disp_qty_df[col] = disp_qty_df[col].apply(lambda v: "-" if (pd.isna(v) or v in ["", None, "-"]) else (f"{int(v):,}원" if isinstance(v, (int, float)) else str(v)))
                else:
                    disp_qty_df[col] = disp_qty_df[col].apply(lambda v: "-" if (pd.isna(v) or v in ["", None, "-"]) else (f"{int(v):,}" if isinstance(v, (int, float)) else str(v)))

            st.dataframe(
                disp_qty_df.style.apply(highlight_qty_rows, axis=1),
                hide_index=True,
                use_container_width=True,
                height=520
            )

            # 하단 배송비 정산 상세 요약 expander
            with st.expander("🚚 일자별 배송비 & 무료배송 정산 상세 내역 (지표별 통계 확인)", expanded=False):
                shipping_view = daily_shipping_df.copy()
                shipping_view.index = range(1, len(shipping_view) + 1)
                st.dataframe(
                    shipping_view.style.format({
                        "총_배송건수": "{:,}건",
                        "총_주문건수": "{:,}건",
                        "총_구매수량": "{:,}개",
                        "총_주문금액": "{:,}원",
                        "무료배송_건수": "{:,}건",
                        "유료배송_건수": "{:,}건",
                        "제주도서_건수": "{:,}건",
                        "기본배송비_합계": "{:,}원",
                        "제주배송비_합계": "{:,}원",
                        "총_배송비_합계": "{:,}원"
                    }),
                    use_container_width=True
                )

        with res_t2:
            st.subheader(f"🏢 {active_brand_name} 전달용 최종 발주서")
            st.caption("사은품 출고 행과 이벤트 대상자가 설정한 맞춤 하이라이트 배경색으로 표시된 최종 제출 양식입니다.")
            
            # 1. 사은품/이벤트 대상자 합산 정산 요약표 (순번 + 주문번호 + 수령인 + 총수량 + 총구매금액 + 지급사은품 ⭐)
            event_order_ids = set()
            event_summary_rows = []

            grouped_orders = processed_df.groupby("order_id", sort=False)
            ev_seq = 1
            for oid, group in grouped_orders:
                all_tags = []
                all_gifts = []
                for _, r in group.iterrows():
                    if r.get("event_tags"):
                        all_tags.extend(r["event_tags"])
                    if r.get("gifts_to_add"):
                        for g in r["gifts_to_add"]:
                            g_desc = f"{g.get('product_name','')} ({g.get('quantity',1)}개)"
                            if g_desc not in all_gifts:
                                all_gifts.append(g_desc)

                if all_tags or all_gifts:
                    event_order_ids.add(oid)
                    first_row = group.iloc[0]
                    tot_qty = group["quantity"].sum()
                    tot_price = group["price"].sum()
                    event_summary_rows.append({
                        "순번": ev_seq,
                        "주문번호": oid,
                        "수령인": first_row.get("receiver_name", ""),
                        "연락처": first_row.get("receiver_phone", ""),
                        "총 구매수량": f"{tot_qty:,}개",
                        "총 구매금액": f"{tot_price:,}원",
                        "지급 사은품": ", ".join(all_gifts) if all_gifts else "사은품 없음 (색인만 적용)",
                        "적용 프로모션": ", ".join(sorted(list(set(all_tags)))) if all_tags else "-"
                    })
                    ev_seq += 1

            if event_summary_rows:
                ev_sum_df = pd.DataFrame(event_summary_rows)
                st.markdown(f"##### 🎯 사은품 & 이벤트 당첨 대상자 정산 요약 (총 {len(ev_sum_df)}명)")
                st.caption("각 대상자의 **순번, 주문번호, 수령인, 총 구매수량, 총 구매금액(원), 지급 사은품** 목록입니다.")
                st.dataframe(
                    ev_sum_df,
                    use_container_width=True,
                    hide_index=True
                )

            # 2. 발주서 필터링 (사용자 요청: 3번 배송비 필터처럼 최종발주서 필터링 지원 ⭐)
            n_all_brand = len(brand_df)
            event_row_mask = brand_df["주문번호"].isin(event_order_ids) if "주문번호" in brand_df.columns else pd.Series([False]*len(brand_df))
            if "_row_highlight" in brand_df.columns:
                event_row_mask = event_row_mask | (brand_df["_row_highlight"].isin(["event", "gift"]))

            n_event_brand = int(event_row_mask.sum())
            n_normal_brand = n_all_brand - n_event_brand

            col_bf1, col_bf2 = st.columns([3, 1])
            with col_bf1:
                final_sheet_filter = st.radio(
                    "🔍 최종 발주서 보기 필터",
                    [
                        f"전체 보기 ({n_all_brand}행)",
                        f"🎁 사은품·이벤트 대상 건만 보기 ({n_event_brand}행)",
                        f"📦 일반 주문 건만 보기 ({n_normal_brand}행)"
                    ],
                    horizontal=True,
                    key="final_order_sheet_filter"
                )

            if "사은품·이벤트" in final_sheet_filter:
                filtered_brand_df = brand_df[event_row_mask].copy()
                st.info(f"🎁 **사은품·이벤트 대상 주문:** 총 {len(event_order_ids)}건 ({len(filtered_brand_df)}행)이 표시 중입니다.")
            elif "일반 주문" in final_sheet_filter:
                filtered_brand_df = brand_df[~event_row_mask].copy()
                st.success(f"📦 **일반 주문:** 총 {filtered_brand_df['주문번호'].nunique() if '주문번호' in filtered_brand_df.columns else len(filtered_brand_df)}건 ({len(filtered_brand_df)}행)이 표시 중입니다.")
            else:
                filtered_brand_df = brand_df.copy()

            row_styles = []
            for _, r in filtered_brand_df.iterrows():
                r_hl = r.get("_row_highlight", "")
                r_clr = r.get("_row_color", "")
                r_str = " ".join(str(val) for val in r.values)

                bg_color = ""
                is_bold = False

                if r_hl == "gift" or "[사은품]" in r_str:
                    bg_color = r_clr if r_clr else color_gift
                    is_bold = True
                elif r_hl == "event":
                    bg_color = r_clr if r_clr else color_event

                style_css = ""
                if bg_color:
                    style_css = f"background-color: {bg_color};"
                    if is_bold:
                        style_css += " font-weight: bold;"
                row_styles.append(style_css)

            meta_cols = [c for c in ["_row_highlight", "_row_color"] if c in filtered_brand_df.columns]
            display_brand_df = filtered_brand_df.drop(columns=meta_cols).copy()
            display_brand_df.index = range(1, len(display_brand_df) + 1)

            def style_table_rows(df_sub):
                return pd.DataFrame(
                    [[row_styles[i]] * len(df_sub.columns) for i in range(len(df_sub))],
                    index=df_sub.index,
                    columns=df_sub.columns
                )

            styled_view = display_brand_df.style.apply(style_table_rows, axis=None)

            st.dataframe(
                styled_view,
                use_container_width=True,
                height=420
            )

        with res_t3:
            st.subheader("💾 셀러 통합 마스터 판매 DB (Raw Data)")
            st.caption("배송비 구분(무료/유료), 합포장 합계액, 단가 적용 및 프로모션 색인이 모두 포함된 원본 마스터 DB입니다.")

            # 유료배송/무료배송 대상자 필터 (사용자 요청 반영 ⭐)
            n_total_rows = len(master_raw_df)
            n_paid_rows = len(master_raw_df[master_raw_df["배송비구분"].str.contains("유료배송", na=False)])
            n_free_rows = len(master_raw_df[master_raw_df["배송비구분"].str.contains("무료배송", na=False)])
            n_jeju_rows = len(master_raw_df[master_raw_df["배송비구분"].str.contains("제주", na=False)])

            col_flt1, col_flt2 = st.columns([3, 1])
            with col_flt1:
                raw_filter = st.radio(
                    "🔍 배송비 대상자 필터링",
                    [
                        f"전체 보기 ({n_total_rows}행)",
                        f"🚚 유료배송 대상자만 보기 ({n_paid_rows}행)",
                        f"🟢 무료배송 대상자만 보기 ({n_free_rows}행)",
                        f"🏝️ 제주/도서 대상자만 보기 ({n_jeju_rows}행)"
                    ],
                    horizontal=True,
                    key="raw_db_shipping_filter"
                )

            view_raw_df = master_raw_df.copy()
            if "유료배송" in raw_filter:
                view_raw_df = view_raw_df[view_raw_df["배송비구분"].str.contains("유료배송", na=False)]
                st.info(f"🚚 **유료배송 대상자: 총 {view_raw_df['order_id'].nunique()}건 ({len(view_raw_df)}행)** — 주문금액이 무료배송 기준({shipping_config.get('free_shipping_threshold', 50000):,}원) 미만으로 배송비 3,000원이 부과되는 주문입니다.")
            elif "무료배송" in raw_filter:
                view_raw_df = view_raw_df[view_raw_df["배송비구분"].str.contains("무료배송", na=False)]
                st.success(f"🟢 **무료배송 대상자: 총 {view_raw_df['order_id'].nunique()}건 ({len(view_raw_df)}행)** — 주문금액이 {shipping_config.get('free_shipping_threshold', 50000):,}원 이상인 무료배송 주문입니다.")
            elif "제주" in raw_filter:
                view_raw_df = view_raw_df[view_raw_df["배송비구분"].str.contains("제주", na=False)]

            view_raw_df.index = range(1, len(view_raw_df) + 1)

            def highlight_raw_rows(row):
                ship_type = str(row.get("배송비구분", ""))
                if "유료배송" in ship_type:
                    return ["background-color: #FFF2E6;"] * len(row)  # 부드러운 피치/오렌지
                elif "무료배송" in ship_type:
                    return ["background-color: #F6FFF8;"] * len(row)  # 부드러운 연초록
                return [""] * len(row)

            st.dataframe(
                view_raw_df.style.apply(highlight_raw_rows, axis=1),
                use_container_width=True,
                height=450
            )

            st.markdown("---")
            s5_n1, s5_n2 = st.columns([1, 1])
            with s5_n1:
                if st.button("👈 4단계 이벤트 & 사은품 설정으로 돌아가기", type="secondary", use_container_width=True, key="btn_goto_step4_from5"):
                    st.session_state["active_step_idx"] = 3
                    st.rerun()
            with s5_n2:
                if st.button("🔄 1단계 브랜드 양식 설정으로 돌아가기", type="secondary", use_container_width=True, key="btn_goto_step1_from5"):
                    st.session_state["active_step_idx"] = 0
                    st.rerun()

    else:
        st.info("👆 먼저 2단계에서 셀러 발주서를 업로드하고 3단계에서 상품명을 매핑해주세요.")
        if st.button("👈 2단계 셀러 발주서 업로드로 이동", type="secondary", key="btn_goto_step2_from5_empty"):
            st.session_state["active_step_idx"] = 1
            st.rerun()
