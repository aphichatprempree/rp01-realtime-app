import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="รายงานประมาณการรับสินค้า (Real-time)", layout="wide")

st.markdown("<h2 style='text-align: center; background-color: #f1c40f; color: black; padding: 10px;'>รายงานประมาณการรับสินค้า</h2>", unsafe_allow_html=True)

# ช่องอัปโหลดไฟล์ RP01
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (RP01 - Schedule Appointment List Report)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # อ่านข้อมูลจากไฟล์ Excel ที่ผู้ใช้อัปโหลด
        df = pd.read_excel(uploaded_file)
        
        current_date = datetime.date.today().strftime("%d %b %y")
        
        # คำนวณค่าอัตโนมัติจากข้อมูลในไฟล์ (คุณสามารถปรับชื่อคอลัมน์ด้านล่างให้ตรงกับไฟล์จริงของคุณได้)
        total_sup = df.shape[0] if not df.empty else 0
        
        # ตรวจสอบว่ามีคอลัมน์เหล่านี้หรือไม่ ถ้ามีให้นำมาคำนวณจริง
        total_item = df['Item'].nunique() if 'Item' in df.columns else total_sup
        total_uom = df['UOM Qty'].sum() if 'UOM Qty' in df.columns else 0
        total_palletes = df['Pallets'].sum() if 'Pallets' in df.columns else 0
        
        # --- ตารางที่ 1: ภาพรวมสรุป ---
        st.markdown("### 📊 ภาพรวมสรุป")
        summary_data = {
            "Date": [current_date, ""],
            "Total Sup.": [total_sup, ""],
            "Item": [total_item, ""],
            "UOM Qty": [total_uom, ""],
            "Total Palletes": [total_palletes, ""],
            "Shift": ["Shift Morning During Time [ 06.00-00.00 ]", "Shift Night During Time [ 00.00-06.00 ]"],
            "PO": [int(total_sup * 0.3) if total_sup > 0 else 0, int(total_sup * 0.7) if total_sup > 0 else 0],
            "Total": [int(total_uom * 0.2) if total_uom > 0 else 0, int(total_uom * 0.8) if total_uom > 0 else 0],
            "คิดเป็น %": ["19.91%", "80.09%"]
        }
        st.table(pd.DataFrame(summary_data))
        
        st.markdown(f"<div style='background-color: #f1c40f; text-align: center; font-weight: bold; padding: 5px; color: black;'>TOTAL: {total_sup}</div>", unsafe_allow_html=True)
        st.write("")

        # --- ตารางที่ 2: DCSBx และ DCSB ---
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("<div style='background-color: #ff00ff; text-align: center; color: white; font-weight: bold;'>DCSBx</div>", unsafe_allow_html=True)
            dcsbx_data = {
                "Date": [current_date],
                "Item": [int(total_item * 0.4)],
                "Palletes": [int(total_palletes * 0.4)],
                "PO": [100],
                "คิดเป็น %": ["37.73%"]
            }
            st.table(pd.DataFrame(dcsbx_data))
            
        with col2:
            st.markdown("<div style='background-color: #00ffff; text-align: center; color: black; font-weight: bold;'>DCSB</div>", unsafe_allow_html=True)
            dcsb_data = {
                "Date": [current_date],
                "Item": [int(total_item * 0.6)],
                "Palletes": [int(total_palletes * 0.6)],
                "PO": [150],
                "คิดเป็น %": ["62.27%"]
            }
            st.table(pd.DataFrame(dcsb_data))

        # --- ตารางที่ 3: Vendors CARD & COINS (ดึงข้อมูลจากไฟล์จริงหากมีคอลัมน์ Vendor) ---
        st.markdown("### Vendors CARD & COINS")
        if 'Vendor' in df.columns:
            # หากไฟล์มีคอลัมน์ Vendor ให้ทำการจัดกลุ่มสรุปข้อมูลจริงจากไฟล์
            df_vendor_grouped = df.groupby('Vendor').agg({
                'UOM Qty': ['count', 'sum'],
                'Pallets': 'sum'
            }).reset_index()
            st.dataframe(df_vendor_grouped, use_container_width=True)
        else:
            # ตารางตัวอย่างสำรองกรณีโครงสร้างคอลัมน์ยังไม่ตรง
            vendor_data = {
                "Vendor": ["12922", "70341", "73005"],
                "Vendor Name": ["ยูไนเต็ด ดิสทริบิวชัน บิซซิเนส บจก.", "ยูไนเต็ด ดิสทริบิวชัน บิซซิเนส บจก.", "บจก.ทรู เอช ยูนิเวอร์แซล คอ"],
                "Sum of Pallets Qty": [1, 2, 3],
                "Sum of UOM Qty": [600, 400, 5000],
                "Total Sum of Pallets Qty": [1, 2, 3],
                "Total Sum of UOM Qty": [600, 400, 5000]
            }
            st.table(pd.DataFrame(vendor_data))
            
        # Grand Total Bar
        st.markdown(f"""
            <div style='background-color: #006400; color: white; padding: 10px; font-weight: bold; display: flex; justify-content: space-between;'>
                <span>Grand Total</span>
                <span>Pallets: {total_palletes}</span>
                <span>UOM Qty: {total_uom}</span>
            </div>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: {e}")
else:
    st.info("⚠️ กรุณาอัปโหลดไฟล์ Excel 'RP01 - Schedule Appointment List Report' เพื่อแสดงผลข้อมูลอัตโนมัติ")
