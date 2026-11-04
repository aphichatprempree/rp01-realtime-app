# --- ส่วนตรวจสอบและสรุปยอดพาเลทของงานนอกรอบ (N/A) แบบป้องกัน NameError ---
st.markdown("---")
st.markdown(
    '<div class="sub-title">🔍 ตรวจสอบงานนอกรอบ (เปรียบเทียบ PO กับ RP01)</div>',
    unsafe_allow_html=True,
)

# ช่องอัปโหลดไฟล์ PO เพิ่มเติม
po_uploaded_file = st.file_uploader(
    "📂 อัปโหลดไฟล์ PO (สำหรับเทียบกับ RP01)",
    type=["xlsx", "xls", "csv"],
    key="po_file_uploader_safe",
)

# ประกาศตัวแปรเริ่มต้น
df_na_jobs = pd.DataFrame()

if po_uploaded_file is not None:
  try:
    if po_uploaded_file.name.endswith(".csv"):
      df_po = pd.read_csv(po_uploaded_file)
    else:
      df_po = pd.read_excel(po_uploaded_file)

    st.success("อัปโหลดไฟล์ PO สำเร็จ!")

    # ตรวจสอบว่ามีตัวแปร df (ข้อมูล RP01 หลัก) อยู่ในระบบหรือไม่
    if "df" in globals() and df is not None and not df.empty:
      po_col_name = next(
          (
              c
              for c in df_po.columns
              if "po" in c.lower() or "order" in c.lower()
          ),
          df_po.columns[0],
      )
      main_po_col = next(
          (c for c in df.columns if "po" in c.lower() or "order" in c.lower()),
          None,
      )

      if main_po_col:
        merged_df = pd.merge(
            df_po,
            df[[main_po_col]],
            left_on=po_col_name,
            right_on=main_po_col,
            how="left",
            indicator=True,
        )
        df_na_jobs = merged_df[merged_df["_merge"] == "left_only"].drop(
            columns=["_merge"]
        )
      else:
        st.warning("ไม่พบคอลัมน์ PO ในตารางหลัก RP01")
    else:
      st.warning(
          "⚠️ กรุณาอัปโหลดไฟล์ RP01 (ไฟล์หลัก) ที่แถบเมนูด้านซ้ายให้เรียบร้อยก่อน"
      )
  except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ PO: {e}")

# แสดงผลเฉพาะเมื่อมีข้อมูลงานนอกรอบ
if not df_na_jobs.empty:
  st.subheader(
      f"📋 รายการงานนอกรอบ (N/A / ไม่มีใน RP01): {len(df_na_jobs)} รายการ"
  )

  # ค้นหาคอลัมน์ Total Case Qty (iLPNs) หรือคอลัมน์พาเลท
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
    df_na_jobs[pallet_col_po] = pd.to_numeric(
        df_na_jobs[pallet_col_po], errors="coerce"
    ).fillna(0)
    total_na_pallets = df_na_jobs[pallet_col_po].sum()

    st.metric(
        label="📦 ยอดพาเลทงานนอกรอบรวม (Total Case Qty / iLPNs)",
        value=f"{total_na_pallets:,.2f} พาเลท",
    )
  else:
    st.info(
        "ℹ️ ไม่พบคอลัมน์ Total Case Qty (iLPNs) ในไฟล์ PO เพื่อรวมยอดพาเลท"
    )

  st.dataframe(df_na_jobs, use_container_width=True)

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
elif po_uploaded_file is not None and "df" in globals() and df is not None:
  st.success("✅ ไม่พบรายการงานนอกรอบ ทุก PO ตรงกับรายงาน RP01 ทั้งหมด")
