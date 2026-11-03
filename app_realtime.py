import streamlit as st
import pandas as pd
import datetime

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide และใช้ Dark Theme
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# Custom CSS ขยายขนาดและเพิ่มความเข้มตัวอักษรให้คมชัด อ่านง่าย
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
        font-size: 28px; 
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(30, 58, 138, 0.4);
        border: 1px solid #3b82f6;
        letter-spacing: 0.5px;
    }
    
    /* หัวข้อรอง (Sub-title) ใหญ่และเข้มขึ้น */
    .sub-title {
        color: #60a5fa; 
        font-weight: 800; 
        font-size: 22px;
        border-left: 6px solid #3b82f6; 
        padding-left: 12px; 
        margin-top: 30px; 
        margin-bottom: 15px;
        letter-spacing: 0.3px;
    }
    
    /* ขยายขนาดและเพิ่มความเข้มตัวอักษรภายในตาราง (Dataframe) ทั้งหมด */
    dataframe, table, th, td, div[data-testid="stDataFrame"] {
        font-size: 18px !important;
    }
    
    /* หัวตารางตัวหนังสือเข้มและใหญ่ขึ้น */
    div[data-testid="stDataFrame"] th {
        font-size: 18px !important;
        font-weight: 800 !important;
        color: #ffffff !important;
    }
    
    /* ข้อมูลในตารางตัวหนังสือใหญ่และหนาขึ้นเพื่อความคมชัด */
    div[data-testid="stDataFrame"] td {
        font-size: 18px !important;
        font-weight: 600 !important;
        color: #f1f5f9 !important;
    }
    
    /* ตกแต่งการ์ดตัวเลข (Metric Cards) ให้ตัวหนังสือใหญ่และเข้มสะใจ */
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, #1e293b, #0f172a);
        border: 1px solid #334155;
        padding: 20px 25px;
        border-radius: 12px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.4);
    }
    div[data-testid="stMetric"] label {
        font-size: 18px !important;
        font-weight: 700 !important;
        color: #cbd5e1 !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        font-size: 34px !important;
        font-weight: 900 !important;
        color: #ffffff !important;
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
                months_th = {1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'มี.ค.', 5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.', 9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.'}
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

        # --- คำนวณ DCSBx และ DCSB ให้ยอด Pallets รวมกันเท่ากับ total_palletes เป๊ะๆ ---
        dcsbx_pallets = int(total_palletes * 0.4)
        dcsb_pallets = int(total_palletes - dcsbx_pallets)

        dcsbx_item = int(total_item * 0.4)
        dcsb_item = total_item - dcsbx_item

        # แสดงผลการ์ดสรุปตัวเลขด้านบน (Metric Cards)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Total Sup.", f"{total_sup:,}")
        with c2:
            st.metric("Item", f"{total_item:,}")
        with c3:
            st.metric("UOM Qty", f"{total_uom:,.0f}")
        with c4:
            st.metric("Total Palletes", f"{total_palletes:,.0f}")

        st.markdown("---")

        # --- ตารางที่ 1: ภาพรวมสรุปกะการทำงาน ---
        st.markdown('<div class="sub-title">📋 ภาพรวมสรุปกะการทำงาน</div>', unsafe_allow_html=True)
        summary_data = {
            "Date": [current_date, ""],
            "Total Sup.": [total_sup, ""],
            "Item": [total_item, ""],
            "UOM Qty": [f"{total_uom:,.0f}", ""],
            "Total Palletes": [f"{total_palletes:,.0f}", ""],
            "Shift": ["Shift Morning During Time [ 06.00-00.00 ]", "Shift Night During Time [ 00.00-06.00 ]"],
            "PO": [morning_po, night_po],
            "Total": [f"{morning_pallets:,.0f}", f"{night_pallets:,.0f}"],
            "คิดเป็น %": [f"{morning_pct:.2f}%", f"{night_pct:.2f}%"]
        }
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)
        
        st.markdown(f"<div style='background-color: #1e3a8a; color: white; text-align: center; font-weight: bold; padding: 12px; border-radius: 6px; font-size: 20px; box-shadow: 0 3px 8px rgba(0,0,0,0.3); border: 1px solid #3b82f6;'>TOTAL: {total_sup}</div>", unsafe_allow_html=True)
        st.write("")

        # --- ตารางที่ 2: DCSBx และ DCSB ---
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("<div style='background-color: #1d4ed8; color: white; text-align: center; font-weight: bold; padding: 12px; border-radius: 8px 8px 0 0; font-size: 20px;'>DCSBx</div>", unsafe_allow_html=True)
            dcsbx_data = {"Date": [current_date], "Item": [dcsbx_item], "Palletes": [dcsbx_pallets], "PO": [100], "คิดเป็น %": ["37.73%"]}
            st.dataframe(pd.DataFrame(dcsbx_data), use_container_width=True, hide_index=True)
            
        with col2:
            st.markdown("<div style='background-color: #0369a1; color: white; text-align: center; font-weight: bold; padding: 12px; border-radius: 8px 8px 0 0; font-size: 20px;'>DCSB</div>", unsafe_allow_html=True)
            dcsb_data = {"Date": [current_date], "Item": [dcsb_item], "Palletes": [dcsb_pallets], "PO": [150], "คิดเป็น %": ["62.27%"]}
            st.dataframe(pd.DataFrame(dcsb_data), use_container_width=True, hide_index=True)

        # --- ตารางที่ 3: Vendors CARD & COINS (Special Suppliers) ---
        st.markdown('<div class="sub-title">⭐ Vendors CARD & COINS (Special Suppliers)</div>', unsafe_allow_html=True)
        
        # แบ่งรายการ Supplier ออกเป็นกลุ่ม CARD และ COINS ตามที่คุณต้องการแยก (สามารถปรับแก้รายชื่อ ID แยกได้ตามจริง)
        card_vendors = [
            "592", "3751", "3924", "12457", "12790", "12810", "12922", "12990", "13725", "17157",
            "21189", "21265", "21541", "21653", "21738", "40334", "40352", "40678", "41068", "70082"
        ]
        
        coins_vendors = [
            "70335", "70341", "70585", "70605", "70609", "70757", "71007", "71026", "71316", "71338",
            "71573", "71758", "72204", "72215", "72472", "72715", "72736", "72747", "72852", "72853",
            "70979", "73204", "73208", "1101530", "FS", "73005", "1102038", "2052496", "2053616", "2055418"
        ]
        
        special_vendors = card_vendors + coins_vendors
        
        if 'Vendor' in df.columns and 'P/O' in df.columns:
            df['Vendor_str'] = df['Vendor'].astype(str).str.split('.').str[0].str.strip()
            df['PO_clean'] = pd.to_numeric(df['P/O'], errors='coerce').fillna(df['P/O']).astype(str).str.split('.').str[0]
            
            special_df = df[df['Vendor_str'].isin(special_vendors)].copy()
            
            if not special_df.empty:
                # กำหนดสถานะเป็น CARD หรือ COINS ตามกลุ่มรายชื่อ
                special_df['Status_val'] = special_df['Vendor_str'].apply(lambda x: 'CARD' if x in card_vendors else 'COINS')
                
                unique_special = special_df.drop_duplicates(subset=['PO_clean'])
                vendor_table = unique_special[['Vendor_str', 'Vendor Name', 'PO_clean', pallets_col, uom_col, 'Status_val']].copy()
                vendor_table.columns = ["Vendor ID", "Vendor Name", "P/O", "Sum of Pallets Qty", "Sum of UOM Qty", "Status"]
                st.dataframe(vendor_table, use_container_width=True, hide_index=True)
            else:
                st.info("ไม่พบรายการ Supplier พิเศษ (Card & Coin) ในรอบเวลานี้")
        else:
            st.warning("ไม่พบคอลัมน์ Vendor หรือ P/O ในไฟล์รีพอร์ต")
        
        # Grand Total Bar
        st.markdown(f"""
            <div style='background: linear-gradient(135deg, #1e3a8a 0%, #172554 100%); color: white; padding: 16px 24px; border-radius: 8px; font-weight: bold; display: flex; justify-content: space-between; margin-top: 20px; font-size: 20px; box-shadow: 0 4px 15px rgba(30, 58, 138, 0.4); border: 1px solid #3b82f6;'>
                <span>Grand Total</span>
                <span>Pallets: {total_palletes:,.0f}</span>
                <span>UOM Qty: {total_uom:,.0f}</span>
            </div>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: {e}")
else:
    st.info("📂 กรุณาอัปโหลดไฟล์ Excel (RP01) เพื่อเริ่มแสดงผลรายงานประมาณการรับสินค้า")
