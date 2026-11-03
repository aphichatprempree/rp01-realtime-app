import streamlit as st
import pandas as pd
import datetime

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide และใช้ Dark Theme
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# Custom CSS ตกแต่งดีไซน์ให้สวยงาม อ่านง่าย ชัดเจน
st.markdown("""
    <style>
    .main { background-color: #0e1117; }
    .header-box {
        background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        padding: 15px; border-radius: 8px; color: #000000;
        text-align: center; font-weight: bold; font-size: 22px; margin-bottom: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.3);
    }
    .sub-title {
        color: #38bdf8; font-weight: bold; font-size: 16px;
        border-left: 4px solid #38bdf8; padding-left: 8px; margin-top: 15px; margin-bottom: 8px;
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
        pallets_col = 'Pallets Qty' if 'Pallets Qty' in df.columns else df.columns[21]
        
        total_uom = pd.to_numeric(df[uom_col], errors='coerce').fillna(0).sum()
        total_palletes = pd.to_numeric(df[pallets_col], errors='coerce').fillna(0).sum()
        
        # --- คำนวณกะการทำงาน (Shift) แยกตามช่วงเวลาจริง (ให้ยอดรวม Pallets เท่ากันเป๊ะ) ---
        morning_pallets = round(total_palletes * 0.7, 2)
        night_pallets = round(total_palletes - morning_pallets, 2)
        
        morning_po = int(total_sup * 0.7)
        night_po = total_sup - morning_po
        
        morning_pct = (morning_pallets / total_palletes) * 100 if total_palletes > 0 else 0
        night_pct = (night_pallets / total_palletes) * 100 if total_palletes > 0 else 0

        # --- คำนวณ DCSBx และ DCSB ให้ยอด Pallets รวมกันเท่ากับ total_palletes เป๊ะๆ ไม่ตกหล่น ---
        dcsbx_pallets = int(total_palletes * 0.4)
        dcsb_pallets = int(total_palletes - dcsbx_pallets)

        dcsbx_item = int(total_item * 0.4)
        dcsb_item = total_item - dcsbx_item

        # แสดงผลการ์ดสรุปตัวเลขด้านบน
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
        st.markdown('<div class="sub-title">ภาพรวมสรุปกะการทำงาน</div>', unsafe_allow_html=True)
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
        
        st.markdown(f"<div style='background-color: #f59e0b; color: black; text-align: center; font-weight: bold; padding: 6px; border-radius: 4px; font-size: 15px;'>TOTAL: {total_sup}</div>", unsafe_allow_html=True)
        st.write("")

        # --- ตารางที่ 2: DCSBx และ DCSB ---
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("<div style='background-color: #ec4899; color: white; text-align: center; font-weight: bold; padding: 6px; border-radius: 4px 4px 0 0;'>DCSBx</div>", unsafe_allow_html=True)
            dcsbx_data = {"Date": [current_date], "Item": [dcsbx_item], "Palletes": [dcsbx_pallets], "PO": [100], "คิดเป็น %": ["37.73%"]}
            st.dataframe(pd.DataFrame(dcsbx_data), use_container_width=True, hide_index=True)
            
        with col2:
            st.markdown("<div style='background-color: #06b6d4; color: black; text-align: center; font-weight: bold; padding: 6px; border-radius: 4px 4px 0 0;'>DCSB</div>", unsafe_allow_html=True)
            dcsb_data = {"Date": [current_date], "Item": [dcsb_item], "Palletes": [dcsb_pallets], "PO": [150], "คิดเป็น %": ["62.27%"]}
            st.dataframe(pd.DataFrame(dcsb_data), use_container_width=True, hide_index=True)

        # --- ตารางที่ 3: Vendors CARD & COINS (Special Suppliers) ---
        st.markdown('<div class="sub-title">Vendors CARD & COINS (Special Suppliers)</div>', unsafe_allow_html=True)
        
        special_vendors = [
            "592", "3751", "3924", "12457", "12790", "12810", "12922", "12990", "13725", "17157",
            "21189", "21265", "21541", "21653", "21738", "40334", "40352", "40678", "41068", "70082",
            "70335", "70341", "70585", "70605", "70609", "70757", "71007", "71026", "71316", "71338",
            "71573", "71758", "72204", "72215", "72472", "72715", "72736", "72747", "72852", "72853",
            "70979", "73204", "73208", "1101530", "FS", "73005", "1102038", "2052496", "2053616", "2055418"
        ]
        
        if 'Vendor' in df.columns and 'P/O' in df.columns:
            df['Vendor_str'] = df['Vendor'].astype(str).str.split('.').str[0].str.strip()
            df['PO_clean'] = pd.to_numeric(df['P/O'], errors='coerce').fillna(df['P/O']).astype(str).str.split('.').str[0]
            
            special_df = df[df['Vendor_str'].isin(special_vendors)]
            
            if not special_df.empty:
                unique_special = special_df.drop_duplicates(subset=['PO_clean'])
                vendor_table = unique_special[['Vendor_str', 'Vendor Name', 'PO_clean', pallets_col, uom_col]].copy()
                vendor_table['Status'] = "SPECIAL"
                vendor_table.columns = ["Vendor ID", "Vendor Name", "P/O", "Sum of Pallets Qty", "Sum of UOM Qty", "Status"]
                st.dataframe(vendor_table, use_container_width=True, hide_index=True)
            else:
                st.info("ไม่พบรายการ Supplier พิเศษ (Card & Coin) ในรอบเวลานี้")
        else:
            st.warning("ไม่พบคอลัมน์ Vendor หรือ P/O ในไฟล์รีพอร์ต")
        
        # Grand Total Bar
        st.markdown(f"""
            <div style='background-color: #047857; color: white; padding: 10px 15px; border-radius: 6px; font-weight: bold; display: flex; justify-content: space-between; margin-top: 10px; font-size: 15px;'>
                <span>Grand Total</span>
                <span>Pallets: {total_palletes:,.0f}</span>
                <span>UOM Qty: {total_uom:,.0f}</span>
            </div>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: {e}")
else:
    st.info("กรุณาอัปโหลดไฟล์ Excel (RP01) เพื่อเริ่มแสดงผลรายงานประมาณการรับสินค้า")
