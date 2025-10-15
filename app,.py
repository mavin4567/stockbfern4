# StockbFern — Smart PDF→Excel Updater (Thai UI)
# อ่าน PDF → Match SKU (B) + สี (C) → Sum ลงคอลัมน์ D
# มี Fuzzy Matching + Confidence % + Flag + Log + Unknown SKU

import io
import re
import pandas as pd
import numpy as np
import streamlit as st
import pdfplumber
from rapidfuzz import process, fuzz

st.set_page_config(page_title="StockbFern — Smart PDF→Excel", page_icon="📦", layout="wide")

# ---------- CONFIG ----------
COLOR_KEYWORDS = [
    "ขาว","ดำ","แดง","น้ำเงิน","ฟ้า","เขียว","เหลือง","ชมพู","ม่วง","ส้ม","เทา","เงิน","ทอง",
    "ทองด้าน","ดำด้าน","ลายจุด","ลายดอก","ใส","ขุ่น","ครีม","น้ำตาล","ฟ้าอ่อน","ฟ้าเข้ม",
    "แดงเข้ม","แดงสด","โรสโกลด์","ทองชมพู",
    "white","black","red","blue","green","yellow","pink","purple","orange",
    "grey","gray","silver","gold","rose gold","matte black","clear"
]

# ---------- UTILITIES ----------
def norm_text(s: str) -> str:
    if s is None: return ""
    s = str(s).strip()
    return re.sub(r"\s+", " ", s)

def norm_key(s: str) -> str:
    if s is None: return ""
    s = str(s).lower()
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[-_./\\]+","", s)
    return s

def detect_color(full_text: str):
    t = full_text.lower()
    hits = [kw for kw in COLOR_KEYWORDS if kw.lower() in t]
    if not hits: return ""
    hits.sort(key=len, reverse=True)
    return hits[0]

def extract_blocks_from_pdf(file_obj):
    """อ่านไฟล์ PDF แบบ line-by-line"""
    records = []
    with pdfplumber.open(file_obj) as doc:
        for page in doc.pages:
            txt = page.extract_text() or ""
            if not txt.strip():
                continue
            lines = [l.strip() for l in txt.splitlines() if l.strip()]
            for i, line in enumerate(lines):
                qty = None
                if re.fullmatch(r"\d+", line):  # ถ้าเจอบรรทัดเป็นตัวเลขล้วน
                    qty = int(line)
                    if i > 0:
                        full = " ".join(lines[max(0, i-3):i])  # รวมชื่อก่อนหน้า
                        records.append({"full": norm_text(full), "qty": qty})
    return pd.DataFrame(records)

def try_extract_code_from_text(full_text: str, excel_codes_norm_set):
    t = norm_key(full_text)
    for code_norm in excel_codes_norm_set:
        if code_norm and code_norm in t:
            return code_norm
    return ""

# ---------- UI ----------
st.title("📦 StockbFern — Smart PDF→Excel (ไทย)")
st.caption("อ่านชื่อสินค้า (เต็มข้อความ) + จำนวนจาก PDF → จับคู่กับ Excel (B=SKU, C=สี) → รวมผลลงคอลัมน์ D พร้อมแถวใหม่")

with st.sidebar:
    st.header("📁 อัปโหลดไฟล์")
    pdf_files = st.file_uploader("เลือกไฟล์ PDF (หลายไฟล์ได้)", type=["pdf"], accept_multiple_files=True)
    xlsx_file = st.file_uploader("เลือกไฟล์ Excel (B=SKU, C=สี, D=จำนวน)", type=["xlsx"])
    sheet_name = st.text_input("ชื่อชีต (เว้นว่าง = ใช้ชีตแรก)", value="")
    st.markdown("---")
    min_fuzzy = st.slider("เกณฑ์ Fuzzy ที่ต้องตั้ง Flag (%)", 50, 90, 65, 1)
    run_btn = st.button("🚀 เริ่มประมวลผล")

status = st.empty()
status.info("READY • พร้อมรับไฟล์")

# ---------- MAIN ----------
if run_btn:
    if not pdf_files or not xlsx_file:
        st.error("กรุณาอัปโหลดทั้ง PDF และ Excel ก่อนเริ่ม")
        st.stop()

    status.warning("RUNNING • กำลังอ่านข้อมูล...")

    all_parts = []
    for f in pdf_files:
        df = extract_blocks_from_pdf(io.BytesIO(f.read()))
        if not df.empty:
            df["file"] = f.name
            all_parts.append(df)

    if not all_parts:
        st.error("ไม่พบข้อมูลที่อ่านได้จาก PDF")
        st.stop()

    pdf_df = pd.concat(all_parts, ignore_index=True)
    pdf_df["color_guess"] = pdf_df["full"].map(detect_color)

    st.success(f"อ่านข้อมูลจาก PDF สำเร็จ: {len(pdf_df)} รายการ")
    with st.expander("🔍 แสดงข้อมูลจาก PDF"):
        st.dataframe(pdf_df, use_container_width=True, hide_index=True)

    # อ่าน Excel
    xl = pd.ExcelFile(xlsx_file)
    sh = sheet_name if sheet_name and sheet_name in xl.sheet_names else xl.sheet_names[0]
    base = xl.parse(sh, header=None)

    while base.shape[1] < 4:
        base[base.shape[1]] = ""
    base = base.reindex(columns=range(4))

    col_B, col_C, col_D = 1, 2, 3
    base[col_B] = base.iloc[:, col_B].astype(str).fillna("")
    base[col_C] = base.iloc[:, col_C].astype(str).fillna("")
    base[col_D] = pd.to_numeric(base.iloc[:, col_D], errors="coerce").fillna(0).astype(int)
    base["__B_norm"] = base.iloc[:, col_B].map(norm_key)
    base["__C_norm"] = base.iloc[:, col_C].map(norm_key)
    base["__BC_norm"] = base["__B_norm"] + "|" + base["__C_norm"]
    excel_codes_norm = set(base["__B_norm"].unique())

    add_qty = np.zeros(len(base), dtype=int)
    unknown_rows = []
    logs = []

    excel_targets = (base.iloc[:, col_B] + " " + base.iloc[:, col_C]).fillna("").astype(str).tolist()
    excel_targets_norm = [norm_text(x) for x in excel_targets]

    for i, row in pdf_df.iterrows():
        full = row["full"]; qty = int(row["qty"]); color_guess = row["color_guess"]
        code_from_contains = try_extract_code_from_text(full, excel_codes_norm)
        match_index, confidence = -1, 0
        if code_from_contains:
            candidates = base.index[base["__B_norm"] == code_from_contains].tolist()
            if candidates:
                color_norm = norm_key(color_guess)
                same_color = [x for x in candidates if base.loc[x, "__C_norm"] == color_norm]
                if same_color:
                    match_index = same_color[0]; confidence = 100
                else:
                    match_index = candidates[0]; confidence = 90
        if match_index < 0:
            best = process.extractOne(norm_text(full), excel_targets_norm, scorer=fuzz.token_set_ratio)
            if best and best[1] >= 65:
                match_index = best[2]; confidence = best[1]
        if match_index >= 0:
            add_qty[match_index] += qty
            logs.append((full, base.iloc[match_index, col_B], base.iloc[match_index, col_C], confidence))
        else:
            unknown_rows.append({"full": full, "color": color_guess, "qty": qty})
            logs.append((full, "-", "-", 0))

    # update quantities
    base.iloc[:, col_D] = (base.iloc[:, col_D] + add_qty).astype(int)

    # Export Results
    status.success("FINISH • ประมวลผลเสร็จแล้ว ✅")

    # Show logs
    log_df = pd.DataFrame(logs, columns=["PDF_SKU", "Matched_SKU", "Color", "Confidence(%)"])
    with st.expander("🧾 Log / Confidence Report"):
        st.dataframe(log_df, use_container_width=True, hide_index=True)
        st.download_button("⬇️ ดาวน์โหลด Log (.xlsx)", data=log_df.to_excel(index=False, engine="openpyxl"),
                           file_name="log_report.xlsx")

    if unknown_rows:
        st.warning(f"พบสินค้าใหม่ {len(unknown_rows)} รายการ ที่ยังไม่อยู่ใน Excel เดิม")
        unknown_df = pd.DataFrame(unknown_rows)
        st.dataframe(unknown_df, use_container_width=True)
        st.download_button("⬇️ ดาวน์โหลดรายการ SKU ใหม่", data=unknown_df.to_excel(index=False, engine="openpyxl"),
                           file_name="unknown_sku.xlsx")

    # Export Excel
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        out_full = base.drop(columns=[c for c in base.columns if str(c).startswith("__")], errors="ignore")
        out_full.to_excel(xw, index=False, header=False, sheet_name=str(sh)[:31])
        pdf_df.to_excel(xw, index=False, sheet_name="PDF_Extract")
    st.download_button("⬇️ ดาวน์โหลด Excel ผลลัพธ์", data=bio.getvalue(),
                       file_name="stockbfern_output.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
