import io
import math
import os
import pandas as pd
import numpy as np
import msoffcrypto
import streamlit as st

st.set_page_config(page_title="Excel Converter", page_icon="📄")
st.title("📄 Excel Converter (Password Remove + Transform)")

st.markdown("""
**What it does**
1. Removes password from an uploaded `.xlsx` file.
2. In **`hawb`** sheet:
   - If `manufacture_name` (BW) length > 100, keep first half of the string.
   - If `manufacture_address` (BX) length > 225, keep first half.
   - If `STATE` column exists (or unnamed column at position 23), remove it.
   - If there is a column at position 77 (BZ), and its string length > 8, keep first half. *(Optional per your last version)*
3. In **`mawb`** sheet:
   - Set cell **L2** (i.e., row index 0 in pandas for the first row) under column `consignee_id_number` to **`2567704`**.
""")

# ---------- Helpers ----------
def truncate_half_if_over(s, threshold):
    if isinstance(s, str) and len(s) > threshold:
        return s[: len(s) // 2]
    return s

def decrypt_xlsx(uploaded_file, password: str) -> io.BytesIO:
    """Decrypt an uploaded XLSX (BytesIO) with a password and return a BytesIO of the decrypted file."""
    decrypted = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(uploaded_file)
    office_file.load_key(password=password)
    office_file.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted

def safe_get_col_by_name_or_index(df: pd.DataFrame, preferred_name: str, index_fallback: int | None):
    """
    Return the column name to use:
      - If preferred_name exists, use it.
      - Else if index_fallback is within bounds, use df.columns[index_fallback].
      - Else return None.
    """
    if preferred_name in df.columns:
        return preferred_name
    if index_fallback is not None and 0 <= index_fallback < len(df.columns):
        return df.columns[index_fallback]
    return None

# ---------- UI ----------
uploaded = st.file_uploader("Upload password-protected .xlsx", type=["xlsx"])
password = st.text_input("Password", type="password", value="_S8&Dwy2&U")

# Advanced options (optional toggles)
with st.expander("Advanced column settings (optional)"):
    st.markdown("If your columns are unnamed, the app will fall back to column positions.")
    bw_idx = st.number_input("Fallback index for BW (manufacture_name)", min_value=0, value=74, step=1)
    bx_idx = st.number_input("Fallback index for BX (manufacture_address)", min_value=0, value=75, step=1)
    bz_idx = st.number_input("Fallback index for BZ (optional extra truncation)", min_value=0, value=77, step=1)
    unnamed_state_idx = st.number_input("Fallback index for STATE col (Unnamed: 23)", min_value=0, value=23, step=1)

if st.button("Process") and uploaded is not None and password:
    try:
        # 1) Decrypt
        decrypted = decrypt_xlsx(uploaded, password)

        # 2) Load workbook
        xls = pd.ExcelFile(decrypted, engine="openpyxl")

        # ---------- HAWB ----------
        if "hawb" not in xls.sheet_names:
            st.error("Sheet 'hawb' not found in the workbook.")
            st.stop()
        df_hawb = pd.read_excel(xls, sheet_name="hawb")

        # map preferred names with fallbacks
        # Your original sheet used unnamed headers like "Unnamed: 74/75/77". We try by name first, then by index.
        bw_col = safe_get_col_by_name_or_index(df_hawb, "manufacture_name", bw_idx)
        if bw_col is None:
            bw_col = safe_get_col_by_name_or_index(df_hawb, "Unnamed: 74", bw_idx)

        bx_col = safe_get_col_by_name_or_index(df_hawb, "manufacture_address", bx_idx)
        if bx_col is None:
            bx_col = safe_get_col_by_name_or_index(df_hawb, "Unnamed: 75", bx_idx)

        bz_col = safe_get_col_by_name_or_index(df_hawb, "Unnamed: 77", bz_idx)

        # Apply truncations per your rules
        if bw_col in df_hawb.columns:
            df_hawb[bw_col] = df_hawb[bw_col].apply(lambda x: truncate_half_if_over(x, 100))
        if bx_col in df_hawb.columns:
            df_hawb[bx_col] = df_hawb[bx_col].apply(lambda x: truncate_half_if_over(x, 225))
        if bz_col in df_hawb.columns:
            df_hawb[bz_col] = df_hawb[bz_col].apply(lambda x: truncate_half_if_over(x, 8))

        # Remove STATE column if present, or unnamed col at index 23 as in your Colab code
        # Try by explicit name first
        to_drop = []
        if "STATE" in df_hawb.columns:
            to_drop.append("STATE")
        # Then try the unnamed fallback
        unnamed_23_name = None
        if 0 <= unnamed_state_idx < len(df_hawb.columns):
            unnamed_23_name = df_hawb.columns[int(unnamed_state_idx)]
            # only drop if it actually exists and isn't the same as previously added
            if unnamed_23_name not in to_drop and unnamed_23_name in df_hawb.columns:
                # If it's explicitly named "Unnamed: 23", still okay to drop
                if str(unnamed_23_name).startswith("Unnamed:") or unnamed_23_name == "Unnamed: 23":
                    to_drop.append(unnamed_23_name)
                   
         coo_col = safe_get_col_by_name_or_index(df_hawb, "country_of_origin", 63)
         if not coo_col:
             coo_col = safe_get_col_by_name_or_index(df_hawb, "Unnamed: 63", 63)
         
         if coo_col:
             df_hawb[coo_col] = df_hawb[coo_col].astype("string")
             df_hawb[coo_col] = df_hawb[coo_col].replace(r"^\s*$", pd.NA, regex=True)
             df_hawb[coo_col] = df_hawb[coo_col].fillna("CN")

        if to_drop:
            df_hawb.drop(columns=to_drop, inplace=True, errors="ignore")

        # ---------- MAWB ----------
        if "mawb" not in xls.sheet_names:
            st.error("Sheet 'mawb' not found in the workbook.")
            st.stop()
        df_mawb = pd.read_excel(xls, sheet_name="mawb")

        # Set L2 under 'consignee_id_number' -> pandas row 0, column 'consignee_id_number'
        if "consignee_id_number" not in df_mawb.columns:
            st.warning("Column 'consignee_id_number' not found in 'mawb'—will add it.")
            if len(df_mawb) == 0:
                df_mawb = pd.DataFrame({"consignee_id_number": ["2567704"]})
            else:
                # ensure at least 1 row
                if len(df_mawb) < 1:
                    df_mawb.loc[0] = np.nan
                df_mawb.loc[0, "consignee_id_number"] = "2567704"
        else:
            if len(df_mawb) == 0:
                df_mawb.loc[0, "consignee_id_number"] = "2567704"
            else:
                df_mawb.iloc[0, df_mawb.columns.get_loc("consignee_id_number")] = "2567704"

        # 3) Write back to a new Excel in-memory
        out_buf = io.BytesIO()
        with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
            df_hawb.to_excel(writer, sheet_name="hawb", index=False)
            df_mawb.to_excel(writer, sheet_name="mawb", index=False)
        out_buf.seek(0)

        st.success("Done! Download your converted file below.")
        st.download_button(
            label="⬇️ Download converted.xlsx",
            data=out_buf,
            file_name="converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Optional: Preview
        with st.expander("Preview (first 10 rows) – hawb"):
            st.dataframe(df_hawb.head(10))
        with st.expander("Preview (first 10 rows) – mawb"):
            st.dataframe(df_mawb.head(10))

    except Exception as e:
        st.error(f"Processing failed: {e}")
else:
    st.info("Upload a file and enter the password, then click **Process**.")
