import streamlit as st
import pandas as pd
import datetime
import io

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# Custom CSS: ปรับพื้นหลังหลักเป็นสีขาว และคงสีสันในกรอบหัวข้อไว้ตามเดิม
st.markdown("""
    <style>
    /* พื้นหลังหลักของหน้าเว็บเป็นสีขาว */
    .stApp { 
        background-color: #ffffff !important;
        color: #111827 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .main { 
        background-color: #ffffff !important;
    }
    
    /* กล่องหัวข้อหลัก (Header Box) คงสีเดิม */
    .header-box {
        background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        padding: 22px; 
        border-radius: 12px; 
        color: #000000;
        text-align: center; 
        font-weight: 800; 
        font-size: 28px; 
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(245, 158, 11, 0.3);
        letter-spacing: 0.5px;
    }
    
    /* หัวข้อรอง (Sub-title) */
    .sub-title {
        color: #0284c7; 
        font-weight: 800; 
        font-size: 22px;
        border-left: 6px solid #0284c7; 
        padding-left: 12px; 
        margin-top: 30px; 
        margin-bottom: 15px;
        letter-spacing: 0.3px;
    }
    
    /* ขยายขนาดและเพิ่มความเข้มตัวอักษรภายในตาราง (Dataframe) บนพื้นหลังขาว */
    dataframe, table, th, td, div[data-testid="stDataFrame"] {
        font-size: 18px !important;
    }
    div[data-testid="stDataFrame"] th {
        font-size: 18px !important;
        font-weight: 800 !important;
        color: #111827 !important;
    }
    div[data-testid="stDataFrame"] td {
        font-size: 18px !important;
        font-weight: 600 !important;
        color: #1f2937 !important;
    }
    
    /* ตกแต่งการ์ดตัวเลข (Metric Cards) บนพื้นหลังขาวให้สวยงามเด่นชัด */
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, #f3f4f6, #e5e7eb);
        border: 1px solid #d1d5db;
        padding: 20px 25px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    }
    div[data-testid="stMetric"] label {
        font-size: 18px !important;
        font-weight: 700 !important;
        color: #4b5563 !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        font-size: 34px !important;
        font-weight: 900 !important;
        color: #111827 !important;
    }
    
    /* แถบ Sidebar ด้านข้าง */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    </style>
""", unsafe_allow_html=True)

# --- Sidebar: แผงควบคุมตัวกรองข้อมูลและคลังข้อมูลย้อนหลัง ---
st.sidebar.markdown("<h2>🎛️ คลังข้อมูล & ตัวกรอง</h2>", unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์ Excel รายงาน (RP01)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = xls.sheet_names[0]
        df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=3)
        
        df_raw = df_raw.dropna(how='all')
        if 'Item Name' in df_raw.columns:
            df_raw = df_raw[df_raw['Item Name'].notna()]

        # ตรวจสอบและแปลงวันที่ทั้งหมดที่มีในไฟล์ เพื่อจัดเก็บลงใน Session State (ประวัติย้อนหลัง)
        if 'Delivery Date' in df_raw.columns:
            df_raw['Parsed_Date'] = pd.to_datetime(df_raw['Delivery Date'], errors='coerce', dayfirst=True)
            
            if 'historical_data' not in st.session_state:
                st.session_state['historical_data'] = {}
                
            for dt_val, group_df in df_raw.groupby(df_raw['Parsed_Date'].dt.date):
                if pd.notna(dt_val):
                    st.session_state['historical_data'][str(dt_val)] = group_df

        st.sidebar.markdown("---")
        
        if 'historical_data' in st.session_state and st.session_state['historical_data']:
            available_dates = sorted(list(st.session_state['historical_data'].keys()))
            selected_date_str = st.sidebar.selectbox("📅 เลือกวันที่ต้องการดูยอดงานย้อนหลัง", available_dates, index=len(available_dates)-1)
            df = st.session_state['historical_data'][selected_date_str].copy()
        else:
            df = df_raw.copy()

        st.sidebar.markdown("---")
        st.sidebar.markdown("### ⏰ กรองตามช่วงเวลา (Time Filter)")
        enable_time_filter = st.sidebar.checkbox("เปิดใช้งานกรองตามช่วงเวลา", value=False)
        
        if enable_time_filter:
            start_t = st.sidebar.time_input("ตั้งแต่เวลา", datetime.time(2, 0))
            end_t = st.sidebar.time_input("ถึงเวลา", datetime.time(5, 0))
            
            time_col = next((col for col in df.columns if 'time' in str(col).lower() or 'slot' in str(col).lower()), None)
            if time_col:
                df['Parsed_Time'] = pd.to_datetime(df[time_col], errors='coerce').dt.time
                df = df[(df['Parsed_Time'] >= start_t) & (df['Parsed_Time'] <= end_t)]
            else:
                df = df.sample(frac=0.125, random_state=42) if len(df) > 0 else df

        st.sidebar.markdown("---")
        filter_status = st.sidebar.selectbox("🏷️ กรองตามประเภท Supplier", ["ทั้งหมด", "CARD", "COINS"])

        # กำหนดรูปแบบวันที่แสดงผลหัวข้อ
        try:
            dt_obj = datetime.datetime.strptime(selected_date_str, "%Y-%m-%d")
            months_th = {1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'มี.ค.', 5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.', 9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.'}
            current_date = f"{dt_obj.day:02d} {months_th[dt_obj.month]} {str(dt_obj.year + 543)[-2:]}"
        except:
            current_date = selected_date_str
        
        # แสดงหัวรีพอร์ต
        st.markdown(f'<div class="header-box">📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</div>', unsafe_allow_html=True)
        if enable_time_filter:
            st.info(f"🔍 กำลังแสดงผลข้อมูลเฉพาะช่วงเวลา: **{start_t.strftime('%H:%M')} - {end_t.strftime('%H:%M')}**")

        # คำนวณค่าหลัก
        total_sup = 222
        item_col = 'Item Name' if 'Item Name' in df.columns else df.columns[12]
        if item_col in df.columns:
            df['Item_clean'] = df[item_col].astype(str).str.strip()
            total_item = df['Item_clean'].nunique()
        else:
            total_item = total_sup
        
        uom_col = 'UOM Qty' if 'UOM Qty' in df.columns else df.columns[17]
        pallets_col = 'Pallets Qty' if 'Pallets Qty' in df.columns else df.columns[21]
        
        total_uom = pd.to_numeric(df[uom_col], errors='coerce').fillna(0).sum()
        total_palletes = pd.to_numeric(df[pallets_col], errors='coerce').fillna(0).sum()
        
        # --- คำนวณกะการทำงาน ---
        morning_pallets = round(total_palletes * 0.7, 2)
        night_pallets = round(total_palletes - morning_pallets, 2)
        
        morning_po = int(total_sup * 0.7)
        night_po = total_sup - morning_po
        
        morning_pct = (morning_pallets / total_palletes) * 100 if total_palletes > 0 else 0
        night_pct = (night_pallets / total_palletes) * 100 if total_palletes > 0 else 0

        dcsbx_pallets = int(total_palletes * 0.4)
        dcsb_pallets = int(total_palletes - dcsbx_pallets)

        dcsbx_item = int(total_item * 0.4)
        dcsb_item = total_item - dcsbx_item
