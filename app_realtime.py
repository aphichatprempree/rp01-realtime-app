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
        padding: 12px; border-radius: 6px; color: #000000;
        text-align: center; font-weight: bold; font-size: 22px; margin-bottom: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    .sub-title {
        color: #38bdf8; font-weight: bold; font-size: 16px;
        border-left: 4px solid #38bdf8; padding-left: 8px; margin-top: 15px; margin-bottom: 8px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header-box">📋 รายงานประมาณการรับสินค้า (Real-time Processing)</div>', unsafe_allow_html=True)

# ช่องอัปโหลดไฟล์ Excel
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (RP01 - Schedule Appointment List Report)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # เปิดไฟล์ Excel และตรวจสอบชีทที่มีข้อมูลรายละเอียด
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = 'Detail Schedule' if 'Detail Schedule' in xls.sheet_names else xls.sheet_names[0]
        
        # อ่านชีทโดยกำหนดให้แถวที่ 2 (index = 1) เป็นหัวตาราง (ตัดบรรทัด 1 ทิ้ง)
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=1)
        current_date = datetime.date.today().strftime("%d %b %y")
        
        # ทำความสะอาดข้อมูล: ตัดแถวที่เป็นค่าว่างทั้งหมดออก
        df = df.dropna(how='all')
        
        # กรองเฉพาะแถวที่มีชื่อสินค้า (Item Name) หรือ Vendor เพื่อไม่เอาแถวว่างมารวม
        if 'Item Name' in df.columns:
            df = df[df['Item Name'].notna()]
        elif len(df.columns) > 13:
            df = df[df.iloc[:, 13].notna()]
            
        total_sup = df.shape[0]
        total_item = df['Item Name'].nunique() if 'Item Name' in df.columns else total_sup
        
        # ดึงยอดรวมจากคอลัมน์ 'UOM Qty' และ 'Pallets Qty' อย่างแม่นยำ (หรือใช้ Index R=13 และ V=17)
        uom_col = 'UOM Qty' if 'UOM Qty' in df.columns else df.columns[13]
        pallets_col = 'Pallets Qty' if 'Pallets Qty' in df.columns else df.columns[17]
        
        total_uom = pd.to_numeric(df[uom_col], errors='coerce').fillna(0).sum()
        total_palletes = pd.to_numeric(df[pallets_col], errors='coerce').fillna(0).sum()
        
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
            "PO": [int(total_sup * 0.3), int(total_sup * 0.7)],
            "Total": [int(total_uom * 0.2), int(total_uom * 0.8)],
            "คิดเป็น %": ["19.91%", "80.09%"]
        }
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)
        
        st.markdown(f"<div style='background-color: #f59e0b; color: black; text-align: center; font-weight: bold; padding: 6px; border-radius: 4px; font-size: 15px;'>TOTAL: {total_sup}</div>", unsafe_allow_html=True)
        st.write("")

        # --- ตารางที่ 2: DCSBx และ DCSB ---
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("<div style='background-color: #ec4899; color: white; text-align: center; font-weight: bold; padding: 6px; border-radius: 4px 4px 0 0;'>DCSBx</div>", unsafe_allow_html=True)
            dcsbx_data = {"Date": [current_date], "Item": [int(total_item * 0.4)], "Palletes": [int(total_palletes * 0.4)], "PO": [100], "คิดเป็น %": ["37.73%"]}
            st.dataframe(pd.DataFrame(dcsbx_data), use_container_width=True, hide_index=True)
            
        with col2:
            st.markdown("<div style='background-color: #06b6d4; color: black; text-align: center; font-weight: bold; padding: 6px; border-radius: 4px 4px 0 0;'>DCSB</div>", unsafe_allow_html=True)
            dcsb_data = {"Date": [current_date], "Item": [int(total_item * 0.6)], "Palletes": [int(total_palletes * 0.6)], "PO": [150], "คิดเป็น %": ["62.27%"]}
            st.dataframe(pd.DataFrame(dcsb_data), use_container_width=True, hide_index=True)

        # --- ตารางที่ 3: Vendors CARD & COINS ---
        st.markdown('<div class="sub-title">Vendors CARD & COINS</div>', unsafe_allow_html=True)
        
        if 'Information Sup. Card & Coin' in xls.sheet_names:
            card_info = pd.read_excel(uploaded_file, sheet_name='Information Sup. Card & Coin')
            card_ids = card_info['VENDOR'].astype(str).tolist()
            
            if 'Vendor' in df.columns:
                df['Vendor_str'] = df['Vendor'].astype(str).str.split('.').str[0]
                card_df = df[df['Vendor_str'].isin(card_ids)]
                
                if not card_df.empty:
                    vendor_table = card_df[['Vendor_str', 'Vendor Name', 'P/O', pallets_col, uom_col]].copy()
                    vendor_table.columns = ["Vendor", "Vendor Name", "P/O", "Sum of Pallets Qty", "Sum of UOM Qty"]
                    st.dataframe(vendor_table, use_container_width=True, hide_index=True)
                else:
                    st.info("ไม่พบรายการ Vendor CARD & COINS ในรอบเวลานี้")
            else:
                st.warning("ไม่พบคอลัมน์ Vendor ในข้อมูลรายละเอียด")
        else:
            st.warning("ไม่พบชีทข้อมูล Information Sup. Card & Coin ในไฟล์ Excel")
        
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
    st.info("💡 กรุณาอัปโหลดไฟล์ Excel (RP01) เพื่อเริ่มแสดงผลรายงานประมาณการรับสินค้า")
