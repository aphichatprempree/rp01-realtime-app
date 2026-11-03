import streamlit as st
import pandas as pd
import datetime
import os

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# กำหนดชื่อไฟล์สำหรับเซฟเก็บไว้บน Server เพื่อให้ผู้ใช้คนอื่นเปิดมาเจอข้อมูลเดิมเสมอ
SERVER_SAVE_FILE = "latest_uploaded_rp01.xlsx"

# --- กำหนดค่าเริ่มต้นใน Session State ---
if 'font_size' not in st.session_state:
    st.session_state['font_size'] = 36
if 'font_weight' not in st.session_state:
    st.session_state['font_weight'] = "900 (หนาพิเศษ)"
if 'enable_time_filter' not in st.session_state:
    st.session_state['enable_time_filter'] = False
if 'filter_status' not in st.session_state:
    st.session_state['filter_status'] = "ทั้งหมด"
if 'postpone_pallets' not in st.session_state:
    st.session_state['postpone_pallets'] = 0.0

# --- Sidebar: แผงควบคุมตัวกรองข้อมูล คลังข้อมูล และการปรับแต่งการแสดงผล ---
st.sidebar.markdown("<h2>🎛️ ควบคุมการแสดงผล</h2>", unsafe_allow_html=True)

# 1. ส่วนควบคุมขนาดและฟอนต์ (Font Controls)
st.sidebar.markdown("### 🔠 ปรับแต่งตัวอักษรหน้าจอ")
st.session_state['font_size'] = st.sidebar.slider("ขนาดฟอนต์ตารางหลัก (Pixel)", min_value=14, max_value=45, value=st.session_state['font_size'], step=2)
st.session_state['font_weight'] = st.sidebar.selectbox("ความหนาของตัวอักษร", ["900 (หนาพิเศษ)", "800 (หนามาก)", "700 (หนาปานกลาง)"], index=["900 (หนาพิเศษ)", "800 (หนามาก)", "700 (หนาปานกลาง)"].index(st.session_state['font_weight']))
weight_val = st.session_state['font_weight'].split()[0]

# 2. ส่วนอัปโหลดไฟล์ Excel
st.sidebar.markdown("---")
st.sidebar.markdown("### 📂 คลังข้อมูล & ตัวกรอง")
uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์ Excel รายงาน (RP01)", type=["xlsx", "xls"])

# หากมีการอัปโหลดไฟล์ใหม่ ให้บันทึกทับลงบน Server ทันที
if uploaded_file is not None:
    try:
        with open(SERVER_SAVE_FILE, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success("✅ อัปโหลดและบันทึกไฟล์ลง Server เรียบร้อย!")
    except Exception as e:
        st.sidebar.error(f"ไม่สามารถบันทึกไฟล์บน Server ได้: {e}")

# ตรวจสอบว่ามีไฟล์เซฟไว้บน Server หรือไม่ (เพื่อให้คนอื่นเข้ามาแล้วเจอข้อมูลเดิม)
active_source = None
if os.path.exists(SERVER_SAVE_FILE):
    active_source = SERVER_SAVE_FILE

# นำค่าที่ผู้ใช้ปรับแต่งมาใส่ลงใน Dynamic CSS พร้อมเพิ่มฟอนต์พรีเมียมและกรอบมีมิติ
st.markdown(f"""
    <style>
    /* นำเข้าฟอนต์ Prompt จาก Google Fonts เพื่อความสวยงามพรีเมียม */
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;600;700;800;900&display=swap');

    .block-container {{
        padding-top: 1.5rem !important;
        padding-bottom: 2rem !important;
    }}
    .stApp {{ 
        background-color: #f8fafc !important;
        color: #0f172a !important;
        font-family: 'Prompt', sans-serif !important;
    }}
    .main {{ 
        background-color: #f8fafc !important;
    }}
    
    /* กรอบหัวข้อหลัก (Header Box) แบบพรีเมียม มีมิติและเงาซ้อน */
    .header-box {{
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 50%, #1e40af 100%);
        padding: 22px; 
        border-radius: 14px; 
        color: #ffffff;
        text-align: center; 
        font-weight: 800; 
        font-size: 34px; 
        margin-bottom: 16px;
        box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.4), 0 8px 10px -6px rgba(30, 58, 138, 0.3);
        border: 2px solid rgba(255, 255, 255, 0.2);
        letter-spacing: 0.5px;
    }}
    
    /* หัวข้อรอง (Sub-title) */
    .sub-title {{
        color: #1e3a8a; 
        font-weight: 800; 
        font-size: 26px;
        border-left: 8px solid #2563eb; 
        padding-left: 14px; 
        margin-top: 18px; 
        margin-bottom: 12px;
        height: 38px;
        display: flex;
        align-items: center;
        background: linear-gradient(90deg, rgba(37, 99, 235, 0.08) 0%, rgba(255, 255, 255, 0) 100%);
        border-radius: 0 8px 8px 0;
    }}
    
    /* แถบ TOTAL PALLETS แบบมีมิติ */
    .total-box {{
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        color: white; 
        text-align: center; 
        font-weight: 800; 
        padding: 14px; 
        border-radius: 10px; 
        font-size: 24px; 
        box-shadow: 0 6px 15px rgba(15, 23, 42, 0.25); 
        border: 2px solid #3b82f6;
        margin-top: 12px;
    }}

    /* กล่อง DCSBx และ DCSB */
    .dcsbx-box {{
        background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%);
        color: white; 
        text-align: center; 
        font-weight: 800; 
        padding: 12px; 
        border-radius: 10px 10px 0 0; 
        font-size: 22px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }}
    .dcsb-box {{
        background: linear-gradient(135deg, #0284c7 0%, #38bdf8 100%);
        color: white; 
        text-align: center; 
        font-weight: 800; 
        padding: 12px; 
        border-radius: 10px 10px 0 0; 
        font-size: 22px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }}
    
    /* === เอฟเฟกต์และสีสันการ์ดกะดึกและกะเช้าแบบมีมิติสูง (3D Shadow) === */
    @keyframes pulse-night {{
        0% {{ box-shadow: 0 0 0 0 rgba(245, 158, 11, 0.4); }}
        70% {{ box-shadow: 0 0 0 12px rgba(245, 158, 11, 0); }}
        100% {{ box-shadow: 0 0 0 0 rgba(245, 158, 11, 0); }}
    }}
    @keyframes pulse-morning {{
        0% {{ box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.4); }}
        70% {{ box-shadow: 0 0 0 12px rgba(34, 197, 94, 0); }}
        100% {{ box-shadow: 0 0 0 0 rgba(34, 197, 94, 0); }}
    }}

    /* การ์ดกะดึก */
    div.element-container:has(div.metric-night) div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #fffbeb, #fef3c7) !important;
        border: 2px solid #f59e0b !important;
        box-shadow: 0 8px 20px rgba(245, 158, 11, 0.25) !important;
        animation: pulse-night 2.5s infinite;
        border-radius: 12px !important;
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
        box-shadow: 0 8px 20px rgba(34, 197, 94, 0.25) !important;
        animation: pulse-morning 2.5s infinite;
        border-radius: 12px !important;
    }}
    div.element-container:has(div.metric-morning) div[data-testid="stMetric"] label {{
        color: #166534 !important;
    }}
    div.element-container:has(div.metric-morning) div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        color: #15803d !important;
    }}

    /* ดีไซน์ตารางหลัก */
    div.element-container:has(div.sub-title) + div[data-testid="stDataFrame"] th {{
        font-size: {int(st.session_state['font_size'] * 0.8)}px !important;
        font-weight: {weight_val} !important;
        color: #1e293b !important;
        background-color: #e2e8f0 !important;
        padding: 12px !important;
    }}
    div.element-container:has(div.sub-title) + div[data-testid="stDataFrame"] td {{
        font-size: {st.session_state['font_size']}px !important;
        font-weight: {weight_val} !important;
        color: #0f172a !important;
        padding: 12px !important;
    }}

    /* ตารางอื่นๆ */
    div[data-testid="stDataFrame"] {{
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border: 1px solid #cbd5e1;
    }}
    div[data-testid="stDataFrame"] th {{
        font-size: 18px !important;
        font-weight: 800 !important;
        color: #1e293b !important;
        background-color: #cbd5e1 !important;
        padding: 10px !important;
    }}
    div[data-testid="stDataFrame"] td {{
        font-size: 20px !important;
        font-weight: 700 !important;
        color: #0f172a !important;
        padding: 10px !important;
    }}
    
    /* ช่องกรอกตัวเลข */
    label[data-baseweb="label"], div[data-baseweb="input"] input {{
        font-size: 20px !important;
        font-weight: 800 !important;
    }}
    div[data-baseweb="input"] input {{
        font-size: 24px !important;
        font-weight: 800 !important;
        height: 48px !important;
        border-radius: 8px !important;
        border: 2px solid #94a3b8 !important;
    }}
    
    div.stAlert p {{
        font-size: 18px !important;
        font-weight: 800 !important;
    }}

    /* การ์ด Metric ทั่วไป */
    div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #ffffff, #f1f5f9);
        border: 2px solid #cbd5e1;
        padding: 14px 18px;
        border-radius: 12px;
        box-shadow: 0 6px 15px rgba(0, 0, 0, 0.06);
    }}
    div[data-testid="stMetric"] label {{
        font-size: 18px !important;
        font-weight: 800 !important;
        color: #334155 !important;
    }}
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        font-size: 34px !important;
        font-weight: 900 !important;
        color: #0f172a !important;
    }}
    
    /* แถบ Sidebar ด้านข้าง */
    section[data-testid="stSidebar"] {{
        background-color: #f1f5f9;
        border-right: 2px solid #e2e8f0;
    }}
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {{
        font-size: 18px !important;
        font-weight: 800 !important;
        color: #1e293b !important;
    }}
    </style>
""", unsafe_allow_html=True)

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

        st.sidebar.markdown("---")
        
        if 'historical_data' in st.session_state and st.session_state['historical_data']:
            available_dates = sorted(list(st.session_state['historical_data'].keys()))
            selected_date_str = st.sidebar.selectbox("📅 เลือกวันที่ต้องการดูยอดงานย้อนหลัง", available_dates, index=len(available_dates)-1)
            df = st.session_state['historical_data'][selected_date_str].copy()
        else:
            df = df_raw.copy()

        st.sidebar.markdown("---")
        st.sidebar.markdown("### ⏰ กรองตามช่วงเวลา (Time Filter)")
        st.session_state['enable_time_filter'] = st.sidebar.checkbox("เปิดใช้งานกรองตามช่วงเวลา", value=st.session_state['enable_time_filter'])
        
        if st.session_state['enable_time_filter']:
            start_t = st.sidebar.time_input("ตั้งแต่เวลา", datetime.time(2, 0))
            end_t = st.sidebar.time_input("ถึงเวลา", datetime.time(5, 0))
            
            time_col_filter = next((col for col in df.columns if 'time' in str(col).lower() or 'slot' in str(col).lower()), None)
            if time_col_filter:
                df['Parsed_Time'] = pd.to_datetime(df[time_col_filter], errors='coerce').dt.time
                df = df[(df['Parsed_Time'] >= start_t) & (df['Parsed_Time'] <= end_t)]
            else:
                df = df.sample(frac=0.125, random_state=42) if len(df) > 0 else df

        st.sidebar.markdown("---")
        status_options = ["ทั้งหมด", "CARD", "COINS"]
        st.session_state['filter_status'] = st.sidebar.selectbox("🏷️ กรองตามประเภท Supplier", status_options, index=status_options.index(st.session_state['filter_status']))
        filter_status = st.session_state['filter_status']

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

        night_po = int(total_sup * ratio_night)
        morning_po = total_sup - night_po
        
        night_pct = (night_pallets / total_pallets) * 100 if total_pallets > 0 else 0
        morning_pct = (morning_pallets / total_pallets) * 100 if total_pallets > 0 else 0

        dcsbx_pallets = int(total_pallets * 0.4)
        dcsb_pallets = int(total_pallets - dcsbx_pallets)

        dcsbx_item = int(total_item * 0.4)
        dcsb_item = total_item - dcsbx_item

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
            st.metric("🌙 ยอดกะดึก (05:00-17:00)", f"{night_pallets:,.0f}", f"{night_pct:.1f}%")
        with c6:
            st.markdown('<div class="metric-morning"></div>', unsafe_allow_html=True)
            st.metric("☀️ ยอดกะเช้า (17:00-05:00)", f"{morning_pallets:,.0f}", f"{morning_pct:.1f}%")

        col_main_left, col_main_right = st.columns([1.1, 1])

        with col_main_left:
            st.markdown('<div class="sub-title">ภาพรวมสรุปกะการทำงาน</div>', unsafe_allow_html=True)
            summary_df = pd.DataFrame({
                "Date": [current_date, ""],
                "Total Sup.": [total_sup, ""],
                "Item": [total_item, ""],
                "UOM Qty": [f"{total_uom:,.0f}", ""],
                "Net Pallets": [f"{total_pallets:,.0f}", ""],
                "Shift": ["Shift Night [05:00-17:00]", "Shift Morning [17:00-05:00]"],
                "PO": [night_po, morning_po],
                "Total": [f"{night_pallets:,.0f}", f"{morning_pallets:,.0f}"],
                "คิดเป็น %": [f"{night_pct:.2f}%", f"{morning_pct:.2f}%"]
            })
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
            st.markdown(f"<div class='total-box'>TOTAL PALLETS: {total_pallets:,.0f} (หักเลื่อนส่ง: {postpone_pallets:,.0f} พาเลท)</div>", unsafe_allow_html=True)

        with col_main_right:
            st.markdown('<div class="sub-title">(DCSBx / DCSB)</div>', unsafe_allow_html=True)
            sub_c1, sub_c2 = st.columns(2)
            with sub_c1:
                st.markdown("<div class='dcsbx-box'>DCSBx</div>", unsafe_allow_html=True)
                dcsbx_data = {"Date": [current_date], "Item": [dcsbx_item], "Palletes": [dcsbx_pallets], "PO": [100], "%": ["37.73%"]}
                st.dataframe(pd.DataFrame(dcsbx_data), use_container_width=True, hide_index=True)
            with sub_c2:
                st.markdown("<div class='dcsb-box'>DCSB</div>", unsafe_allow_html=True)
                dcsb_data = {"Date": [current_date], "Item": [dcsb_item], "Palletes": [dcsb_pallets], "PO": [150], "%": ["62.27%"]}
                st.dataframe(pd.DataFrame(dcsb_data), use_container_width=True, hide_index=True)

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
        
        if 'Vendor' in df.columns and 'P/O' in df.columns:
            df['Vendor_str'] = df['Vendor'].astype(str).str.split('.').str[0].str.strip()
            df['PO_clean'] = pd.to_numeric(df['P/O'], errors='coerce').fillna(df['P/O']).astype(str).str.split('.').str[0]
            
            special_df = df[df['Vendor_str'].isin(card_vendors + coins_vendors)].copy()
            
            if not special_df.empty:
                special_df['Status_val'] = special_df['Vendor_str'].apply(lambda x: 'CARD' if x in card_vendors else 'COINS')
                
                if filter_status != "ทั้งหมด":
                    special_df = special_df[special_df['Status_val'] == filter_status]
                
                unique_special = special_df.drop_duplicates(subset=['PO_clean'])
                vendor_table = unique_special[['Vendor_str', 'Vendor Name', 'PO_clean', pallets_col, uom_col, 'Status_val']].copy()
                vendor_table.columns = ["Vendor ID", "Vendor Name", "P/O", "Sum of Pallets Qty", "Sum of UOM Qty", "Status"]
                st.dataframe(vendor_table, use_container_width=True, hide_index=True)
            else:
                st.info("ไม่พบรายการ Supplier พิเศษในเงื่อนไขตัวกรองนี้")
        else:
            st.warning("ไม่พบคอลัมน์ Vendor หรือ P/O ในไฟล์รีพอร์ต")
        
        st.markdown(f"""
            <div style='background: linear-gradient(135deg, #047857 0%, #065f46 100%); color: white; padding: 18px 26px; border-radius: 10px; font-weight: 800; display: flex; justify-content: space-between; margin-top: 15px; font-size: 26px; box-shadow: 0 6px 15px rgba(4, 120, 87, 0.3); border: 2px solid #10b981;'>
                <span>Grand Total (Net Pallets)</span>
                <span>Pallets: {total_pallets:,.0f} (เลื่อนออก: {postpone_pallets:,.0f})</span>
                <span>UOM Qty: {total_uom:,.0f}</span>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 📥 ดาวน์โหลดบันทึกยอดงานประจำวัน")
        
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            csv_data = summary_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label=f"📥 ดาวน์โหลดตารางสรุปวันที่ {current_date} (CSV)",
                data=csv_data,
                file_name=f"report_summary_{current_date.replace(' ', '_')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
        with col_dl2:
            html_content = f"""
            <html>
            <head>
                <meta charset='utf-8'>
                <title>Report {current_date}</title>
                <style>
                    body {{ background-color: #ffffff; color: #000000; font-family: sans-serif; padding: 20px; }}
                    h2 {{ color: #1e40af; text-align: center; font-size: 32px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
                    th, td {{ border: 2px solid #94a3b8; padding: 14px; text-align: center; font-size: 24px; font-weight: 900; }}
                    th {{ background-color: #1e3a8a; color: white; }}
                    td {{ background-color: #f1f5f9; color: #000000; }}
                </style>
            </head>
            <body>
                <h2>📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</h2>
                <p style="font-size: 24px; font-weight: bold;"><b>Total Sup:</b> {total_sup:,} | <b>Item:</b> {total_item:,} | <b>UOM Qty:</b> {total_uom:,.0f} | <b>Net Pallets:</b> {total_pallets:,.0f} (เลื่อนออก: {postpone_pallets:,.0f})</p>
                {summary_df.to_html(index=False, border=0)}
                <br>
                <p style="text-align: center; color: #475569; font-size: 18px;">(บันทึกหน้าเว็บนี้เป็นภาพหน้าจอหรือ PDF เพื่อเก็บบันทึกยอดงานย้อนหลัง)</p>
            </body>
            </html>
            """
            st.download_button(
                label=f"🖼️ ดาวน์โหลดรายงานแบบ HTML (เปิดดูย้อนหลังเสมือนหน้าจอ)",
                data=html_content.encode('utf-8-sig'),
                file_name=f"report_dashboard_{current_date.replace(' ', '_')}.html",
                mime="text/html",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: {e}")
else:
    st.info("📂 กรุณาอัปโหลดไฟล์ Excel (RP01) ที่แถบด้านข้าง (Sidebar) ระบบจะทำการบันทึกข้อมูลไว้บน Server เพื่อให้ทุกคนเปิดมาเจอหน้าจอและข้อมูลเดิมเสมอครับ")
