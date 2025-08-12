import io
import re
import numpy as np
import pandas as pd
import msoffcrypto
import streamlit as st
from openpyxl import load_workbook

# ======================
# Streamlit UI
# ======================
st.set_page_config(page_title="Excel Converter (Decrypt + Transform)", page_icon="📄")
st.title("📄 Excel Converter (Password Remove + Transform)")

st.markdown("""
**This app will:**
1) Decrypt an uploaded **.xlsx** using the password you provide  
2) On **hawb**: treat row 2 as headers and process data from row 3; apply your requested transforms  
3) On **mawb**: set **L2** (under `consignee_id_number`) to **2567704**  
4) Provide a clean workbook for download
""")

default_pw = "_S8&Dwy2&U"
password = st.text_input("File password", value=default_pw, type="password")
uploaded = st.file_uploader("Upload password-protected .xlsx", type=["xlsx"])
run_btn = st.button("Convert")

# ======================
# Helpers
# ======================
def decrypt_xlsx(file_bytes: bytes, password: str) -> bytes:
    """Return decrypted xlsx as bytes."""
    fin = io.BytesIO(file_bytes)
    fout = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(fin)
    office_file.load_key(password=password)
    office_file.decrypt(fout)
    return fout.getvalue()

def normalize_map(cols):
    """Map normalized (lower, trimmed) -> original column names."""
    mapping = {}
    for c in cols:
        key = str(c).strip().lower()
        mapping.setdefault(key, c)
    return mapping

def find_col(cols_map, name):
    """Find original column by case-insensitive, trimmed match."""
    return cols_map.get(str(name).strip().lower(), None)

def keep_half_if_over(s, limit):
    """If len(s) > limit, keep first half (floor on odd); else return original."""
    if pd.isna(s):
        return s
    s = str(s)
    return s[: len(s) // 2] if len(s) > limit else s

def enforce_zip_6_digits(x):
    """If not exactly 6 digits, set to '123456'."""
    if pd.isna(x):
        return "123456"
    s = str(x).strip()
    return s if re.fullmatch(r"\d{6}", s) else "123456"

def ensure_at_least_one_visible_and_active(wb):
    """OpenPyXL requires at least one visible sheet and an active index."""
    for ws in wb.worksheets:
        ws.sheet_state = "visible"
    wb.active = 0

# ======================
# Main
# ======================
if run_btn:
    if not uploaded:
        st.warning("Please upload a .xlsx file first.")
        st.stop()
    if not password:
        st.warning("Please enter the password.")
        st.stop()

    try:
        with st.spinner("Decrypting..."):
            decrypted = decrypt_xlsx(uploaded.read(), password)

        # Load workbook once to edit mawb and to preserve other sheets
        wb = load_workbook(io.BytesIO(decrypted))

        # ---- mawb: set L2 (under consignee_id_number) to 2567704 ----
        if "mawb" in wb.sheetnames:
            ws = wb["mawb"]
            ws["L2"].value = "2567704"
        else:
            st.warning("Sheet 'mawb' not found. Skipping mawb update.")

        # ---- hawb: two header rows; row 2 = headers, data starts row 3 ----
        if "hawb" not in wb.sheetnames:
            st.error("Sheet 'hawb' not found in workbook.")
            st.stop()

        raw_hawb = pd.read_excel(io.BytesIO(decrypted), sheet_name="hawb", header=None, dtype=str)
        if raw_hawb.shape[0] < 2:
            st.error("The 'hawb' sheet does not contain at least two header rows.")
            st.stop()

        # Build the dataframe using second row as header
        new_cols = raw_hawb.iloc[1].astype(str).tolist()
        hawb_df = raw_hawb.iloc[2:].copy()  # data from 3rd row
        hawb_df.columns = new_cols
        hawb_df.reset_index(drop=True, inplace=True)

        # Column mapping tolerant to case/space
        colmap = normalize_map(hawb_df.columns)

        # 1) Half-length trimming
        col_manu_name = find_col(colmap, "manufacture_name")
        if col_manu_name in hawb_df.columns:
            hawb_df[col_manu_name] = hawb_df[col_manu_name].apply(lambda x: keep_half_if_over(x, 100))

        col_manu_addr = find_col(colmap, "manufacture_address")
        if col_manu_addr in hawb
