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
        if col_manu_addr in hawb_df.columns:
            hawb_df[col_manu_addr] = hawb_df[col_manu_addr].apply(lambda x: keep_half_if_over(x, 225))

        col_manu_state = find_col(colmap, "manufacture_state")
        if col_manu_state in hawb_df.columns:
            hawb_df[col_manu_state] = hawb_df[col_manu_state].apply(lambda x: keep_half_if_over(x, 8))

        # 2) Set country columns to "CN"
        for cname in ["country_of_origin", "manufacture_country"]:
            c = find_col(colmap, cname)
            if c in hawb_df.columns:
                hawb_df[c] = "CN"

        # 3) Zip: if not exactly 6 digits, set to "123456"
        col_zip_exact = find_col(colmap, "manufacture_zip_code ")
        col_zip_trim = find_col(colmap, "manufacture_zip_code")
        col_zip = col_zip_exact if (col_zip_exact in hawb_df.columns) else col_zip_trim
        if col_zip in hawb_df.columns:
            hawb_df[col_zip] = hawb_df[col_zip].apply(enforce_zip_6_digits)

        # 4) Drop recipient_state entirely
        col_recipient_state = find_col(colmap, "recipient_state")
        if col_recipient_state in hawb_df.columns:
            hawb_df.drop(columns=[col_recipient_state], inplace=True)

        # ---- Write back with visibility/active-sheet guards ----
        ensure_at_least_one_visible_and_active(wb)  # before any save

        # Save current edited wb (with mawb updated) to bytes
        base = io.BytesIO()
        wb.save(base)

        # Now open with pandas' ExcelWriter and replace hawb
        out_io = io.BytesIO()
        with pd.ExcelWriter(out_io, engine="openpyxl") as writer:
            writer.book = load_workbook(io.BytesIO(base.getvalue()))
            writer.sheets = {ws.title: ws for ws in writer.book.worksheets}

            # If hawb exists, delete then write fresh
            if "hawb" in writer.book.sheetnames:
                del writer.book["hawb"]

            # If deleting hawb leaves no visible sheets, create a temp
            if not any(ws.sheet_state == "visible" for ws in writer.book.worksheets):
                writer.book.create_sheet("TempVisible")
            writer.book.active = 0

            # Write new hawb
            hawb_df.to_excel(writer, sheet_name="hawb", index=False)

            # Clean up temp if created
            if "TempVisible" in writer.book.sheetnames:
                del writer.book["TempVisible"]

            # Final safety: ensure visibility and set hawb active
            for ws in writer.book.worksheets:
                ws.sheet_state = "visible"
            writer.book.active = max(0, writer.book.sheetnames.index("hawb"))

            writer.save()

        st.success("Conversion complete!")
        st.download_button(
            "Download converted file",
            data=out_io.getvalue(),
            file_name="converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("What was applied"):
            st.write("""
- Decrypted the workbook
- **hawb**: used row 2 as headers; processed from row 3
  - If `manufacture_name` length > 100 → kept first half
  - If `manufacture_address` length > 225 → kept first half
  - If `manufacture_state` length > 8 → kept first half
  - Set `country_of_origin` and `manufacture_country` to `"CN"`
  - If `manufacture_zip_code` (with or without trailing space) not exactly 6 digits → set to `"123456"`
  - Dropped `recipient_state`
- **mawb**: set **L2** to `2567704`
""")

    except msoffcrypto.exceptions.DecryptionError:
        st.error("Decryption failed. Please verify the password.")
    except Exception as e:
        st.exception(e)
