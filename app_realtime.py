# --- ส่วนเพิ่มเติม: สรุปยอดพาเลทของงานนอกรอบ (N/A) ---
if not df_na_jobs.empty:
  st.subheader(
      f"📋 รายการงานนอกรอบ (N/A / ไม่มีใน RP01): {len(df_na_jobs)} รายการ"
  )

  # ค้นหาคอลัมน์ Total Case Qty (iLPNs) หรือคอลัมน์พาเลทในไฟล์ PO
  pallet_col_po = next(
      (
          c
          for c in df_na_jobs.columns
          if "total case" in c.lower()
          or "ilpn" in c.lower()
          or "pallet" in c.lower()
          or "qty" in c.lower()
      ),
      None,
  )

  if pallet_col_po:
    # แปลงข้อมูลเป็นตัวเลขเพื่อความถูกต้องในการรวมยอด
    df_na_jobs[pallet_col_po] = pd.to_numeric(
        df_na_jobs[pallet_col_po], errors="coerce"
    ).fillna(0)
    total_na_pallets = df_na_jobs[pallet_col_po].sum()

    # แสดงผลการ์ดสรุปยอดพาเลทงานนอกรอบ
    st.metric(
        label="📦 ยอดพาเลทงานนอกรอบรวม (Total Case Qty / iLPNs)",
        value=f"{total_na_pallets:,.2f} พาเลท",
    )
  else:
    st.info(
        "ℹ️ ไม่พบคอลัมน์ Total Case Qty (iLPNs) ในไฟล์ PO เพื่อรวมยอดพาเลท"
    )

  # แสดงตารางข้อมูลงานนอกรอบทั้งหมด
  st.dataframe(df_na_jobs, use_container_width=True)

  # ปุ่มดาวน์โหลดรายงานงานนอกรอบเป็น Excel
  from io import BytesIO

  output = BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_na_jobs.to_excel(writer, index=False, sheet_name="NA_Jobs")
  processed_data = output.getvalue()

  st.download_button(
      label="📥 ดาวน์โหลดรายงานงานนอกรอบ (Excel)",
      data=processed_data,
      file_name="NA_Off_Cycle_Jobs.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )
else:
  st.success("✅ ไม่พบรายการงานนอกรอบ ทุก PO ตรงกับรายงาน RP01 ทั้งหมด")
