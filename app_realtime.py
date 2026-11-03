import streamlit as st
import pandas as pd
import datetime
import os

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# กำหนดชื่อไฟล์สำหรับเซฟเก็บไว้บน Server เพื่อให้ผู้ใช้คนอื่นเปิดมาเจอข้อมูลเดิมเสมอ
SERVER_SAVE_FILE = "latest_uploaded_rp01.xlsx"

# --- กำหนดค่าเริ่มต้นใน Session State สำหรับแยกส่วนควบคุม ---
if 'dcs_font_size' not in st.session_state:
    st.session_state['dcs_font_size'] = 22  # ขนาดฟอนต์ตารางบน (DCSBx / DCSB)
if 'vendor_font_size' not in st.session_state:
    st.session_state['vendor_font_size'] = 20  # ขนาดฟอนต์ตารางล่าง (Vendors CARD & COINS)
if 'enable_time_filter' not in st.session_state:
    st.session_state['enable_time_filter'] = False
if 'filter_status' not in st.session_state:
    st.session_state['filter_status'] = "ทั้งหมด"
if 'postpone_pallets' not in st.session_state:
    st.session_state['postpone_pallets'] = 0.0

dcs_size = st.session_state['dcs_font_size']
vendor_size = st.session_state['vendor_font_size']

# นำค่าฟอนต์ที่ปรับแต่งมาใส่ใน Dynamic CSS เพื่อคุมขนาดตัวอักษรตารางทั้งหมด
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;600;700;800;900&display=swap');

    .block-container {{
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }}
    .stApp {{ 
        background-color: #f8fafc !important;
        color: #0f172a !important;
        font-family: 'Prompt', sans-serif !important;
    }}
    .main {{ 
        background-color: #f8fafc !important;
    }}
    
    /* กรอบหัวข้อหลัก (Header Box) */
    .header-box {{
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 50%, #1e40af 100%);
        padding: 22px 24px 20px 24px; 
        border-radius: 12px; 
        color: #ffffff;
        text-align: center; 
        font-weight: 800; 
        font-size: 30px; 
        margin-bottom: 14px;
        box-shadow: 0 6px 15px -4px rgba(37, 99, 235, 0.4);
        border: 2px solid rgba(255, 255, 255, 0.2);
        line-height: 1.5;
    }}
    
    /* หัวข้อรอง (Sub-title) */
    .sub-title {{
        color: #1e40af; 
        font-weight: 800; 
        font-size: 24px;
        border-left: 8px solid #2563eb; 
        padding-left: 12px; 
        margin-top: 10px; 
        margin-bottom: 8px;
        height: 34px;
        display: flex;
        align-items: center;
        background: linear-gradient(90deg, rgba(37, 99, 235, 0.08) 0%, rgba(255, 255, 255, 0) 100%);
        border-radius: 0 8px 8px 0;
    }}

    /* กล่อง DCSBx และ DCSB */
    .dcsbx-box {{
        background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%);
        color: white; 
        text-align: center; 
        font-weight: 800; 
        padding: 10px; 
        border-radius: 8px 8px 0 0; 
        font-size: 22px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }}
    .dcsb-box {{
        background: linear-gradient(135deg, #0284c7 0%, #38bdf8 100%);
        color: white; 
        text-align: center; 
        font-weight: 800; 
        padding: 10px; 
        border-radius: 8px 8px 0 0; 
        font-size: 22px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }}

    /* === ขยายขนาดตัวหนังสือยอด PO (Delta) ในการ์ดให้ใหญ่สะใจ === */
    div[data-testid="stMetric"] div[data-testid="stMetricDelta"] {{
        font-size: 18px !important;
        font-weight: 900 !important;
    }}

    /* === ปรับขนาดฟอนต์ตารางด้านบน (DCSBx / DCSB) ตาม Sidebar === */
    div.element-container:has(iframe) dataframe, div[data-testid="stDataFrame"] {{
        font-size: {dcs_size}px !important;
    }}
    
    /* บังคับขนาดฟอนต์เซลล์และหัวตารางทั้งหมดให้ใหญ่ตามตัวเลื่อน */
    [data-testid="stDataFrame"] div[data-baseweb="tab"] {{
        font-size: {dcs_size}px !important;
    }}

    /* ปรับแต่งฟอนต์ตารางในหน้าเว็บผ่าน CSS สากลสำหรับตาราง Streamlit */
    .stDataFrame [data-testid="StyledMarkupTable"] td, .stDataFrame [data-testid="StyledMarkupTable"] th {{
        font-size: {dcs_size}px !important;
    }}

    /* === เอฟเฟกต์การ์ดกะดึกและกะเช้าแบบมีมิติสูง (3D Shadow) === */
    @keyframes pulse-night {{
        0% {{ box-shadow: 0 0 0 0 rgba(245, 158, 11, 0.4); }}
        70% {{ box-shadow: 0 0 0 10px rgba(245, 158, 11, 0); }}
        100% {{ box-shadow: 0 0 0 0 rgba(245, 158, 11, 0); }}
    }}
    @keyframes pulse-morning {{
        0% {{ box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.4); }}
        70% {{ box-shadow: 0 0 0 10px rgba(34, 197, 94, 0); }}
        100% {{ box-shadow: 0 0 0 0 rgba(34, 197, 94, 0); }}
    }}

    /* การ์ดกะดึก */
    div.element-container:has(div.metric-night) div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #fffbeb, #fef3c7) !important;
        border: 2px solid #f59e0b !important;
        box-shadow: 0 6px 15px rgba(245, 158, 11, 0.25) !important;
        animation: pulse-night 2.5s infinite;
        border-radius: 10px !important;
        padding: 10px 14px !important;
    }}
    div.element-container:has(div.metric-night) div[data-testid="stMetric"] label {{
        color: #92400e !important;
    }}
    div.element-container:has(div.metric-night) div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        color: #b45309 !important;
    }}

    /* การ์ดกะเช้า */
    div.element-container:has(div.metric-morning) div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #f0fdf4, #dcfce7) !important;
        border: 2px solid #22c55e !important;
        box-shadow: 0 6px 15px rgba(34, 197, 94, 0.25) !important;
        animation: pulse-morning 2.5s infinite;
        border-radius: 10px !important;
        padding: 10px 14px !important;
    }}
    div.element-container:has(div.metric-morning) div[data-testid="stMetric"] label {{
        color: #166534 !important;
    }}
    div.element-container:has(div.metric-morning) div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        color: #15803d !important;
    }}
    
    /* ช่องกรอกตัวเลข */
    label[data-baseweb="label"], div[data-baseweb="input"] input {{
        font-size: 18px !important;
        font-weight: 800 !important;
    }}
    div[data-baseweb="input"] input {{
        font-size: 20px !important;
        font-weight: 800 !important;
        height: 40px !important;
        border-radius: 6px !important;
        border: 2px solid #94a3b8 !important;
    }}
    
    div.stAlert p {{
        font-size: 16px !important;
        font-weight: 800 !important;
        padding-top: 2px !important;
        padding-bottom: 2px !important;
    }}

    /* การ์ด Metric ทั่วไป */
    div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #ffffff, #f1f5f9);
        border: 2px solid #cbd5e1;
        padding: 10px 14px;
        border-radius: 10px;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05);
    }}
    div[data-testid="stMetric"] label {{
        font-size: 16px !important;
        font-weight: 800 !important;
        color: #334155 !important;
    }}
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        font-size: 28px !important;
        font-weight: 900 !important;
        color: #0f172a !important;
    }}
    
    /* === ปรับแต่งแถบ Sidebar ด้านข้างใหม่ให้ทันสมัยและมีมิติ === */
    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
        border-right: 2px solid #e2e8f0;
        padding: 1rem 0.8rem;
    }}
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {{
        font-size: 16px !important;
        font-weight: 800 !important;
        color: #1e293b !important;
    }}
    section[data-testid="stSidebar"] .stSelectbox, section[data-testid="stSidebar"] .stSlider, section[data-testid="stSidebar"] .stFileUploader {{
        background: #ffffff;
        padding: 10px;
        border-radius: 8px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03);
        border: 1px solid #e2e8f0;
        margin-bottom: 8px;
    }}

    @media print {{
        section[data-testid="stSidebar"], button, .stDownloadButton {{
            display: none !important;
        }}
        .stApp {{
            background-color: #ffffff !important;
        }}
    }}
    </style>
""", unsafe_allow_html=True)

# --- Sidebar: แผงควบคุมตัวกรองข้อมูล คลังข้อมูล และการปรับแต่งการแสดงผล ---
st.sidebar.markdown("<h2 style='color: #1e40af; border-bottom: 3px solid #2563eb; padding-bottom: 6px;'>🎛️ แผงควบคุมระบบ</h2>", unsafe_allow_html=True)

# 1. ส่วนคลังข้อมูล & อัปโหลดไฟล์
st.sidebar.markdown("### 📂 1. จัดการข้อมูลไฟล์ (RP01)")
uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์ Excel รายงาน", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        with open(SERVER_SAVE_FILE, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success("✅ อัปโหลดและบันทึกไฟล์สำเร็จ!")
    except Exception as e:
        st.sidebar.error(f"ไม่สามารถบันทึกไฟล์: {e}")

active_source = None
if os.path.exists(SERVER_SAVE_FILE):
    active_source = SERVER_SAVE_FILE

st.sidebar.markdown("---")

# 2. ส่วนเลือกวันที่ดูข้อมูลย้อนหลัง
st.sidebar.markdown("### 📅 2. เลือกวันที่ดูยอดงานย้อนหลัง")

if active_source is not None:
    try:
        xls = pd.ExcelFile(active_source)
        target_sheet = xls.sheet_names[0]
        df_raw = pd.read_excel(active_source, sheet_name=target_sheet, header=3)
        
        df_raw = df_raw.dropna(how='all')
        if 'Item Name' in df_raw.columns:
            df_raw = df_raw[df_raw['Item Name'].notna()]

        if 'Delivery Date' in df_raw.columns:
            df_raw['Parsed_Date'] = pd.to_datetime(df_raw['Delivery Date'], errors='coerce', dayfirst=True)
            
            if 'historical_data' not in st.session_state:
                st.session_state['historical_data'] = {}
                
            for dt_val, group_df in df_raw.groupby(df_raw['Parsed_Date'].dt.date):
                if pd.notna(dt_val):
                    st.session_state['historical_data'][str(dt_val)] = group_df

        if 'historical_data' in st.session_state and st.session_state['historical_data']:
            available_dates = sorted(list(st.session_state['historical_data'].keys()))
            selected_date_str = st.sidebar.selectbox("เลือกวันที่ต้องการตรวจสอบ", available_dates, index=len(available_dates)-1)
            df = st.session_state['historical_data'][selected_date_str].copy()
        else:
            df = df_raw.copy()
            selected_date_str = str(datetime.date.today())

    except Exception as e:
        st.sidebar.error(f"อ่านไฟล์ไม่สำเร็จ: {e}")
        df = pd.DataFrame()
        selected_date_str = str(datetime.date.today())
else:
    df = pd.DataFrame()
    selected_date_str = str(datetime.date.today())

st.sidebar.markdown("---")

# 3. ส่วนกรองประเภท Supplier & ช่วงเวลา
st.sidebar.markdown("### 🏷️ 3. ตัวกรองข้อมูลเพิ่มเติม")
status_options = ["ทั้งหมด", "CARD", "COINS"]
st.session_state['filter_status'] = st.sidebar.selectbox("กรองตามประเภท Supplier", status_options, index=status_options.index(st.session_state['filter_status']))
filter_status = st.session_state['filter_status']

st.session_state['enable_time_filter'] = st.sidebar.checkbox("⏰ เปิดใช้งานกรองตามช่วงเวลา", value=st.session_state['enable_time_filter'])
if st.session_state['enable_time_filter']:
    start_t = st.sidebar.time_input("ตั้งแต่เวลา", datetime.time(2, 0))
    end_t = st.sidebar.time_input("ถึงเวลา", datetime.time(5, 0))
    
    time_col_filter = next((col for col in df.columns if 'time' in str(col).lower() or 'slot' in str(col).lower()), None)
    if time_col_filter and not df.empty:
        df['Parsed_Time'] = pd.to_datetime(df[time_col_filter], errors='coerce').dt.time
        df = df[(df['Parsed_Time'] >= start_t) & (df['Parsed_Time'] <= end_t)]

st.sidebar.markdown("---")

# 4. ส่วนปรับขนาดตัวอักษรแยกอิสระแต่ละตาราง
st.sidebar.markdown("### 🔠 4. ปรับขนาดตัวอักษรในตาราง")
st.session_state['dcs_font_size'] = st.sidebar.slider("ขนาดฟอนต์ตารางบน (DCSBx / DCSB)", min_value=16, max_value=50, value=st.session_state['dcs_font_size'], step=2)
st.session_state['vendor_font_size'] = st.sidebar.slider("ขนาดฟอนต์ตารางล่าง (Vendors CARD/COINS)", min_value=16, max_value=45, value=st.session_state['vendor_font_size'], step=2)

# --- ส่วนประมวลผลและแสดงผลหน้าจอหลัก ---
if active_source is not None and not df.empty:
    try:
        try:
            dt_obj = datetime.datetime.strptime(selected_date_str, "%Y-%m-%d")
            months_th = {1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'มี.ค.', 5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.', 9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.'}
            current_date = f"{dt_obj.day:02d} {months_th[dt_obj.month]} {str(dt_obj.year + 543)[-2:]}"
        except:
            current_date = selected_date_str
        
        st.markdown(f'<div class="header-box">📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</div>', unsafe_allow_html=True)
        if st.session_state['enable_time_filter']:
            st.info(f"🔍 กำลังแสดงผลข้อมูลเฉพาะช่วงเวลา: **{start_t.strftime('%H:%M')} - {end_t.strftime('%H:%M')}**")

        st.markdown("---")
        col_postpone1, col_postpone2 = st.columns([1, 2])
        with col_postpone1:
            st.session_state['postpone_pallets'] = st.number_input("⏳ พาเลทที่เลื่อนส่งออก (Pallets):", min_value=0.0, value=float(st.session_state['postpone_pallets']), step=1.0, help="ใส่จำนวนพาเลทที่เลื่อนออก เพื่อหักออกจากยอดรวม")
            postpone_pallets = st.session_state['postpone_pallets']
        with col_postpone2:
            if postpone_pallets > 0:
                st.warning(f"⚠️ มีการเลื่อนส่งงานออกจำนวน **{postpone_pallets:,.0f} พาเลท** ระบบได้ทำการหักลบออกจากยอดรวมสุทธิให้เรียบร้อยแล้ว")
            else:
                st.info("ℹ️ กำลังแสดงยอดพาเลทตามแผนปกติ (ไม่มีงานเลื่อนส่ง)")

        # กำหนดชื่อคอลัมน์และนับจำนวน Unique Values แบบตรงเป๊ะกับไฟล์ต้นฉบับ
        sup_col = 'Vendor Name' if 'Vendor Name' in df.columns else (next((col for col in df.columns if 'vendor' in str(col).lower() or 'supplier' in str(col).lower()), df.columns[4]))
        total_sup = df[sup_col].dropna().nunique() if sup_col in df.columns else len(df)

        item_col = 'Item Name' if 'Item Name' in df.columns else (next((col for col in df.columns if 'item' in str(col).lower()), df.columns[12]))
        total_item = df[item_col].dropna().nunique() if item_col in df.columns else len(df)

        po_col_raw = next((c for c in df.columns if 'p/o' in str(c).lower() or 'po' in str(c).lower()), None)
        if not po_col_raw:
            po_col_raw = 'P/O' if 'P/O' in df.columns else df.columns[1]

        if po_col_raw in df.columns:
            df['PO_clean'] = pd.to_numeric(df[po_col_raw], errors='coerce').fillna(df[po_col_raw]).astype(str).str.split('.').str[0].str.strip()
            total_po = df['PO_clean'].dropna().nunique()
        else:
            total_po = len(df)

        uom_col = 'UOM Qty' if 'UOM Qty' in df.columns else df.columns[17]
        pallets_col = 'Pallets Qty' if 'Pallets Qty' in df.columns else df.columns[21]
        
        total_uom = pd.to_numeric(df[uom_col], errors='coerce').fillna(0).sum()
        raw_pallets = pd.to_numeric(df[pallets_col], errors='coerce').fillna(0).sum()
        
        total_pallets = max(0.0, raw_pallets - postpone_pallets)
        
        # --- กำหนดสัดส่วนตามเป้าหมาย (กะดึก 3879 / กะเช้า 1663) ---
        target_night = 3879
        target_morning = 1663
        total_target = target_night + target_morning
        
        ratio_night = target_night / total_target
        ratio_morning = target_morning / total_target

        night_pallets = round(total_pallets * ratio_night, 2)
        morning_pallets = round(total_pallets - night_pallets, 2)

        # คำนวณยอด PO แยกกะตามสัดส่วนพาเลท
        night_po = round(total_po * ratio_night)
        morning_po = total_po - night_po
        
        night_pct = (night_pallets / total_pallets) * 100 if total_pallets > 0 else 0
        morning_pct = (morning_pallets / total_pallets) * 100 if total_pallets > 0 else 0

        # --- คำนวณแยกตามคอลัมน์ ASRS_ITEM สำหรับ DCSBx และ DCSB ---
        asrs_col = 'ASRS_ITEM' if 'ASRS_ITEM' in df.columns else (next((c for c in df.columns if 'asrs' in str(c).lower()), 'AE'))
        
        if asrs_col in df.columns:
            df['ASRS_match'] = df[asrs_col].astype(str).str.strip().str.upper()
            dcsbx_df_filtered = df[df['ASRS_match'] == 'ASRS_ITEM']
            dcsb_df_filtered = df[df['ASRS_match'] != 'ASRS_ITEM']
        else:
            dcsbx_df_filtered = df.iloc[:int(len(df)*0.4)]
            dcsb_df_filtered = df.iloc[int(len(df)*0.4):]

        # ยอดของ DCSBx (ASRS_ITEM)
        dcsbx_item = dcsbx_df_filtered[item_col].dropna().nunique() if item_col in dcsbx_df_filtered.columns else len(dcsbx_df_filtered)
        dcsbx_pallets = pd.to_numeric(dcsbx_df_filtered[pallets_col], errors='coerce').fillna(0).sum()
        dcsbx_po = dcsbx_df_filtered['PO_clean'].dropna().nunique() if 'PO_clean' in dcsbx_df_filtered.columns else len(dcsbx_df_filtered)
        dcsbx_pct = (dcsbx_pallets / total_pallets) * 100 if total_pallets > 0 else 0

        # ยอดของ DCSB (งานปกติ)
        dcsb_item = dcsb_df_filtered[item_col].dropna().nunique() if item_col in dcsb_df_filtered.columns else len(dcsb_df_filtered)
        dcsb_pallets = pd.to_numeric(dcsb_df_filtered[pallets_col], errors='coerce').fillna(0).sum()
        dcsb_po = dcsb_df_filtered['PO_clean'].dropna().nunique() if 'PO_clean' in dcsb_df_filtered.columns else len(dcsb_df_filtered)
        dcsb_pct = (dcsb_pallets / total_pallets) * 100 if total_pallets > 0 else 0

        # แสดงผลการ์ดสรุปตัวเลขด้านบน
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            st.metric("Total Sup.", f"{total_sup:,}")
        with c2:
            st.metric("Item", f"{total_item:,}")
        with c3:
            st.metric("UOM Qty", f"{total_uom:,.0f}")
        with c4:
            st.metric("Net Pallets", f"{total_pallets:,.0f}", delta=f"-{postpone_pallets:,.0f} เลื่อนออก" if postpone_pallets > 0 else None, delta_color="inverse")
        with c5:
            st.markdown('<div class="metric-night"></div>', unsafe_allow_html=True)
            st.metric("🌙 กะดึก (05:00-17:00)", f"{night_pallets:,.0f}", f"{night_po:,} PO ({night_pct:.1f}%)")
        with c6:
            st.markdown('<div class="metric-morning"></div>', unsafe_allow_html=True)
            st.metric("☀️ กะเช้า (17:00-05:00)", f"{morning_pallets:,.0f}", f"{morning_po:,} PO ({morning_pct:.1f}%)")

        # --- ตาราง DCSBx / DCSB (แยกตาม ASRS_ITEM จริง) ---
        st.markdown('<div class="sub-title">(DCSBx / DCSB)</div>', unsafe_allow_html=True)
        sub_c1, sub_c2 = st.columns(2)
        with sub_c1:
            st.markdown("<div class='dcsbx-box'>DCSBx (ASRS_ITEM)</div>", unsafe_allow_html=True)
            dcsbx_table_df = pd.DataFrame({
                "Date": [current_date], 
                "Item": [dcsbx_item], 
                "Palletes": [dcsbx_pallets], 
                "PO": [dcsbx_po], 
                "%": [f"{dcsbx_pct:.2f}%"]
            })
            st.dataframe(dcsbx_table_df, use_container_width=True, hide_index=True)
        with sub_c2:
            st.markdown("<div class='dcsb-box'>DCSB (Normal)</div>", unsafe_allow_html=True)
            dcsb_table_df = pd.DataFrame({
                "Date": [current_date], 
                "Item": [dcsb_item], 
                "Palletes": [dcsb_pallets], 
                "PO": [dcsb_po], 
                "%": [f"{dcsb_pct:.2f}%"]
            })
            st.dataframe(dcsb_table_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown('<div class="sub-title">Vendors CARD & COINS (Special Suppliers)</div>', unsafe_allow_html=True)
        
        card_vendors = [
            "592", "3751", "3924", "12457", "12790", "12810", "12922", "12990", "13725", "17157",
            "21189", "21265", "21541", "21653", "21738", "40334", "40352", "40678", "41068", "70082"
        ]
        coins_vendors = [
            "70335", "70341", "70585", "70605", "70609", "70757", "71007", "71026", "71316", "71338",
            "71573", "71758", "72204", "72215", "72472", "72715", "72736", "72747", "72852", "72853",
            "70979", "73204", "73208", "1101530", "FS", "73005", "1102038", "2052496", "2053616", "2055418"
        ]
        
        v_col = next((c for c in df.columns if 'vendor' in str(c).lower() and 'name' not in str(c).lower()), None)
        if not v_col:
            v_col = 'Vendor' if 'Vendor' in df.columns else df.columns[3]

        if v_col in df.columns and po_col_raw in df.columns:
            df['Vendor_str'] = df[v_col].astype(str).str.split('.').str[0].str.strip()
            
            special_df = df[df['Vendor_str'].isin(card_vendors + coins_vendors)].copy()
            
            if not special_df.empty:
                special_df['Status_val'] = special_df['Vendor_str'].apply(lambda x: 'CARD' if x in card_vendors else 'COINS')
                
                if filter_status != "ทั้งหมด":
                    special_df = special_df[special_df['Status_val'] == filter_status]
                
                unique_special = special_df.drop_duplicates(subset=['PO_clean'])
                vendor_table = unique_special[['Vendor_str', sup_col, 'PO_clean', pallets_col, uom_col, 'Status_val']].copy()
                vendor_table.columns = ["Vendor ID", "Vendor Name", "P/O", "Sum of Pallets Qty", "Sum of UOM Qty", "Status"]
                
                st.dataframe(vendor_table, use_container_width=True, hide_index=True)
            else:
                st.info("ℹ️ ไม่พบรายการ Supplier พิเศษ (CARD/COINS) ในเงื่อนไขตัวกรองหรือวันที่เลือกนี้")
        else:
            st.warning("⚠️ ไม่พบคอลัมน์ Vendor หรือ P/O ในไฟล์รีพอร์ต")
        
        st.markdown(f"""
            <div style='background: linear-gradient(135deg, #047857 0%, #065f46 100%); color: white; padding: 14px 22px; border-radius: 8px; font-weight: 800; display: flex; justify-content: space-between; margin-top: 10px; font-size: 24px; box-shadow: 0 4px 12px rgba(4, 120, 87, 0.3); border: 2px solid #10b981;'>
                <span>Grand Total (Net Pallets)</span>
                <span>Pallets: {total_pallets:,.0f} (เลื่อนออก: {postpone_pallets:,.0f})</span>
                <span>UOM Qty: {total_uom:,.0f}</span>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 📥 บันทึกและพิมพ์รายงาน")
        
        col_dl1, col_dl2, col_dl3 = st.columns(3)
        
        with col_dl1:
            summary_df_export = pd.DataFrame({
                "Date": [current_date],
                "Total Sup.": [total_sup],
                "Item": [total_item],
                "UOM Qty": [f"{total_uom:,.0f}"],
                "Net Pallets": [f"{total_pallets:,.0f}"]
            })
            csv_data = summary_df_export.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label=f"📥 ดาวน์โหลดตารางสรุป (CSV)",
                data=csv_data,
                file_name=f"report_summary_{current_date.replace(' ', '_')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
        with col_dl2:
            html_filename = f"report_dashboard_{current_date.replace(' ', '_')}.html"
            html_content = f"""<html>
<head>
    <meta charset='utf-8'>
    <title>Report {current_date}</title>
    <style>
        body {{ background-color: #ffffff; color: #000000; font-family: 'Prompt', sans-serif; padding: 20px; }}
        h2 {{ color: #1e40af; text-align: center; font-size: 28px; font-weight: 800; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        th, td {{ border: 2px solid #94a3b8; padding: 10px; text-align: center; font-size: 18px; font-weight: 800; }}
        th {{ background-color: #1e3a8a; color: white; }}
        td {{ background-color: #f1f5f9; color: #000000; }}
    </style>
</head>
<body>
    <h2>📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</h2>
    <p style="font-size: 18px; font-weight: bold; text-align: center;"><b>Total Sup:</b> {total_sup:,} | <b>Item:</b> {total_item:,} | <b>UOM Qty:</b> {total_uom:,.0f} | <b>Net Pallets:</b> {total_pallets:,.0f} (เลื่อนออก: {postpone_pallets:,.0f})</p>
    <br>
    <div style="text-align: center;">
        <button onclick="window.print()" style="background-color: #2563eb; color: white; padding: 12px 24px; font-size: 18px; font-weight: bold; border: none; border-radius: 8px; cursor: pointer;">🖨️ พิมพ์หน้านี้ / บันทึกเป็น PDF</button>
    </div>
</body>
</html>"""
            st.download_button(
                label=f"📄 ดาวน์โหลดไฟล์หน้าเว็บ (HTML)",
                data=html_content.encode('utf-8-sig'),
                file_name=html_filename,
                mime="text/html",
                use_container_width=True
            )

        with col_dl3:
            st.markdown(f"""
                <div style="display: flex; align-items: center; height: 100%;">
                    <button onclick="window.print()" style="width: 100%; background: linear-gradient(135deg, #2563eb, #1e40af); color: white; padding: 10px 16px; font-size: 18px; font-weight: 800; border: none; border-radius: 8px; cursor: pointer; box-shadow: 0 4px 10px rgba(37,99,235,0.3);">
                        🖨️ พิมพ์/บันทึก PDF ทันที
                    </button>
                </div>
            """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: {e}")
else:
    st.info("📂 กรุณาอัปโหลดไฟล์ Excel (RP01) ที่แถบด้านข้าง (Sidebar) ระบบจะทำการบันทึกข้อมูลไว้บน Server เพื่อให้ทุกคนเปิดมาเจอหน้าจอและข้อมูลเดิมเสมอครับ")
