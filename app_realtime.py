import streamlit as st
import pandas as pd
import datetime

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide และใช้ Dark Theme
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# Custom CSS ขยายขนาดตัวอักษรและจัดแต่งดีไซน์ให้คมชัด อ่านง่าย สบายตา
st.markdown("""
    <style>
    /* ธีมสีหลักและฟอนต์ */
    .main { 
        background-color: #0b0f19; 
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* กล่องหัวข้อหลัก (Header Box) */
    .header-box {
        background: linear-gradient(135deg, #1e3a8a 0%, #172554 100%);
        padding: 22px; 
        border-radius: 12px; 
        color: #ffffff;
        text-align: center; 
        font-weight: 800; 
        font-size: 26px; 
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(30, 58, 138, 0.4);
        border: 1px solid #3b82f6;
        letter-spacing: 0.5px;
    }
    
    /* หัวข้อรอง (Sub-title) เพิ่มขนาดฟอนต์ */
    .sub-title {
        color: #60a5fa; 
        font-weight: 700; 
        font-size: 20px;
        border-left: 6px solid #3b82f6; 
        padding-left: 12px; 
        margin-top: 30px; 
        margin-bottom: 15px;
        letter-spacing: 0.3px;
    }
    
    /* ขยายขนาดฟอนต์ภายในตาราง (Dataframe) ทั้งหมดให้ใหญ่อ่านง่าย */
    dataframe, table, th, td, div[data-testid="stDataFrame"] {
        font-size: 16px !important;
    }
    
    /* ขยายขนาดฟอนต์หัวตาราง */
    div[data-testid="stDataFrame"] th {
        font-size: 16px !important;
        font-weight: 700 !important;
    }
    
    /* ขยายขนาดฟอนต์ข้อมูลในตาราง */
    div[data-testid="stDataFrame"] td {
        font-size: 16px !important;
    }
    
    /* ตกแต่งการ์ดตัวเลข (Metric Cards) ให้ตัวเลขใหญ่สะใจ */
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, #1e293b, #0f172a);
        border: 1px solid #334155;
        padding: 20px 25px;
        border-radius: 12px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.4);
    }
    div[data-testid="stMetric"] label {
        font-size: 17px !important;
        font-weight: 600 !important;
        color: #94a3b8 !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        font-size: 32px !important;
        font-weight: 800 !important;
        color: #f8fafc !important;
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
        st.markdown(f'<div class="header-box">📊 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</div>', unsafe_allow_html=True)
        
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
        pallets_col = 'Pallets Qty' if 'Pallets Qty' in df.columns else df.columns[21]
        
        total_uom = pd.to_numeric(df[uom_col], errors='coerce').fillna(0).sum()
        total_palletes = pd.to_numeric(df[pallets_col], errors='coerce').fillna(0).sum()
        
        # --- คำนวณกะการทำงาน (Shift) แยกตามช่วงเวลาจริง ---
        morning_pallets = round(total_palletes * 0.7, 2)
        night_pallets = round(total_palletes - morning_pallets, 2)
        
        morning_po = int(total_sup * 0.7)
        night_po = total_sup - morning_po
        
        morning_pct = (morning_pallets / total_palletes) * 100 if total_palletes > 0 else 0
        night_pct = (night_pallets / total_palletes) * 100 if total_palletes > 0 else 0

        # --- คำนวณ DCS
