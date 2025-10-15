import streamlit as st
import pandas as pd
import pdfplumber
import re
from rapidfuzz import fuzz, process
from io import BytesIO

st.set_page_config(page_title="Omni Picklist OCR", page_icon="📦", layout="wide")

st.title("📦 Omni Picklist OCR & Excel Updater")
st.caption("ระบบรวมยอด SKU จากไฟล์ PDF และอัปเดตไฟล์ Excel template โดยอัตโนมัติ")

# ========== Upload ==========
uploaded_pdfs = st.file_uploader("📁 เลือกไฟล์ PDF Picklist (หลายไฟล์ได้)", type=["pdf"], accept_multiple_files=True)
uploaded_excel = st.file_uploader("📊 เลือกไฟล์ Excel Template", type=["xlsx"])

# ========== Extract Function ==========
def extract_sku_quantity_from_pdf(file):
    """ดึงข้อมูล SKU และจำนวนจาก PDF"""
    data = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split("\n")
            for line in lines:
                # ค้นหา SKU
                sku_match = re.search(r"SKU:\s*([A-Za-z0-9\-\._\s/]+)", line)
                # ค้นหาจำนวน (เลขท้ายบรรทัด)
                qty_match = re.search(r"\b(\d+)$", line.strip())
                if sku_match and qty_match:
                    sku = sku_match.group(1).strip()
                    qty = int(qty_match.group(1))
                    data.append({"SKU": sku, "Qty": qty})
    return pd.DataFrame(data)

def fuzzy_match(sku, sku_list, threshold=80):
    """จับคู่ SKU แบบใกล้เคียง"""
    match = process.extractOne(sku, sku_list, scorer=fuzz.partial_ratio)
    if match and match[1] >= threshold:
        return match[0]
    return None

# ========== Main Process ==========
if uploaded_pdfs and uploaded_excel:
    st.success("✅ โหลดไฟล์สำเร็จ! เริ่มประมวลผล...")
    all_data = pd.DataFrame(columns=["SKU", "Qty"])

    for pdf in uploaded_pdfs:
        df = extract_sku_quantity_from_pdf(pdf)
        all_data = pd.concat([all_data, df], ignore_index=True)

    # รวมยอด SKU
    sku_sum = all_data.groupby("SKU", as_index=False)["Qty"].sum()
    st.subheader("📋 ข้อมูล SKU ที่ดึงจาก PDF")
    st.dataframe(sku_sum, use_container_width=True)

    # อ่านไฟล์ Excel
    template_df = pd.read_excel(uploaded_excel)
    st.subheader("📊 Template เดิม")
    st.dataframe(template_df.head(), use_container_width=True)

    # เตรียม DataFrame สำหรับอัปเดต
    updated_df = template_df.copy()
    updated_df["จำนวน (ใหม่)"] = 0

    for i, row in updated_df.iterrows():
        sku_excel = str(row["SKU"]).strip()
        match_sku = fuzzy_match(sku_excel, sku_sum["SKU"].tolist())
        if match_sku:
            qty = sku_sum.loc[sku_sum["SKU"] == match_sku, "Qty"].sum()
            updated_df.loc[i, "จำนวน (ใหม่)"] = qty

    st.subheader("✅ ผลลัพธ์หลังอัปเดต")
    st.dataframe(updated_df, use_container_width=True)

    # ========== Download ==========
    output = BytesIO()
    updated_df.to_excel(output, index=False)
    st.download_button(
        "⬇️ ดาวน์โหลดไฟล์ Excel ที่อัปเดตแล้ว",
        data=output.getvalue(),
        file_name="Updated_Picklist.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("กรุณาอัปโหลดทั้งไฟล์ PDF และไฟล์ Excel ก่อนเริ่มทำงาน")

st.divider()
st.caption("พัฒนาโดย BLS Automation Lab – ใช้ OCR และ fuzzy matching เพื่อช่วยจับคู่ SKU อัตโนมัติ 🧠")
