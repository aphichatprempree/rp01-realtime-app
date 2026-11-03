import streamlit as st
import pandas as pd
import datetime

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide และใช้ Dark Theme
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# Custom CSS ตกแต่งดีไซน์ให้ดูสวยงาม มีมิติ และพรีเมียมยิ่งขึ้น
st.markdown("""
    <style>
    /* พื้นหลังหลักและฟอนต์ */
    .main { 
        background-color: #0e1117; 
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* กล่องหัวข้อหลัก (Header Box) มีมิติและเงา */
    .header-box {
        background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        padding: 18px; 
        border-radius: 10px; 
        color: #000000;
        text-align: center; 
        font-weight: 800; 
        font-size: 24px; 
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(245, 158, 11, 0.3);
        letter-spacing: 0.5px;
    }
    
    /* หัวข้อรอง (Sub-title) */
    .sub-title {
        color: #38bdf8; 
        font-weight: 700; 
        font-size: 18px;
        border-left: 5px solid #38bdf8; 
        padding-left: 10px; 
        margin-top: 25px; 
        margin-bottom: 12px;
        letter-spacing: 0.3px;
    }
    
    /* ตกแต่ง Metric Cards ด้านบนให้มีมิติ นูนสวยงาม */
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, #1f2937, #111827);
        border: 1px solid #374151;
        padding: 15px 20px;
        border-radius: 12px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(56, 189, 248, 0.2);
        border-color: #38bdf8;
    }
    div[data-testid="stMetric"] label {
        font-size: 15px !important;
        font-weight: 600 !important;
        color: #9ca3af !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        font-size: 28px !important;
        font-weight: 800 !important;
        color: #f3f4f6 !important;
    }
    
    /* ตกแต่งตารางข้อมูล (DataFrame) ให้แสดงผลคมชัด อ่านง่าย */
    dataframe, table {
        border-radius: 8px !important;
        overflow: hidden;
    }
    </style>
""", unsafe_allow_html=True)

# ช่องอัปโหลดไฟล์ Excel
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (RP01 - Schedule Appointment List Report)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # เปิดไฟล์ Excel และอ่านข้อมูลจากแถวหัวตารางจริง (header=3 สำหรับไฟล์รูปแบบใหม่)
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = xls.sheet_names[0]
        
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=3)
        
        # ดึงวันที่จากคอลัมน์ 'Delivery Date' แถวแรกสุดมาอ้างอิง
        raw_date_val = df['Delivery Date'].dropna().iloc[0] if 'Delivery Date' in df.columns and not df['Delivery Date'].dropna().empty else None
        
        if raw_date_val:
            try:
                dt_parsed = pd.to_datetime(raw_date_val, dayfirst=True)
                months_th = {1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'เม.ย.', 5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.', 9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.'}
                current_date = f"{dt_parsed.day:02d} {months_th[dt_parsed.month]} {str(dt_parsed.year + 543)[-2:]}"
            except:
                current_date = str(raw_date_val)
        else:
            current_date = datetime.date.today().strftime("%d %b %y")
        
        # แสดงหัวรีพอร์ตพร้อมวันที่อ้างอิงจาก Delivery Date
        st.markdown(f'<div class="header-box">📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</div>', unsafe_allow_html=True)
        
        # ทำความสะอาดข้อมูล: ตัดแถวที่เป็นค่าว่างทั้งหมดออก
        df = df.dropna(how='all')
        
        # กรองเฉพาะแถวที่มีข้อมูลสินค้าจริง (Item Name)
        if 'Item Name' in df.columns:
            df = df[df['Item Name'].notna()]
            
        # กำหนด Total Sup. เป็น 222 ตามที่คุณต้องการ
        total_sup = 222
        
        # คำนวณ Item แบบ Unique (ไม่ซ้ำกัน) และ Clean ข้อมูลแล้ว
        item_col = 'Item Name' if 'Item Name' in df.columns else df.columns[12]
        if item_col in df.columns:
            df['Item_clean'] = df[item_col].astype(str).str.strip()
            total_item = df['Item_clean'].nunique()
        else:
            total_item = total_sup
        
        # ดึงยอดรวมจากคอลัมน์ UOM Qty และ Pallets Qty อย่างแม่นยำ
        uom_col = 'UOM Qty' if 'UOM Qty' in df.columns else df.columns[17]
        pallets_col = 'Pallets Qty' if 'Pallets Qty' in df.
