import io
import numpy as np
import pandas as pd
import msoffcrypto
import streamlit as st

st.set_page_config(page_title="Excel Converter", page_icon="📄")
st.title("📄 Excel Converter (Password Remove + Transform)")

st.markdown("""
**What this app does**
1. Decrypts an uploaded **.xlsx** using the password you provide.
2. In **`hawb`**:
   - If `manufacture_name` (**BW**) length > 100, keep first half.
   - If `manufacture_address` (**BX**) length > 225, keep first half.
   - If `Unnamed: 77` (approx **BZ**) length > 8, keep first half. *(optional rule)*
   - **Fill missing `country_of_origin` (or `Unnamed: 63`) with `"CN"`** (treats NaN/empty/whitespace as missing).
   - Remove `STATE` column if present; else drop unnamed column at index **23** (e.g., `Unnamed: 23`).
3. In **`mawb`**:
   - Set **L2** (column `consignee_id_number`, row 2 in Excel UI) to **`2567704`**.
4. Download the processed file.
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
    Return a column name to use:
      - If preferred_name exists, use it.
      - Else if index_fallback is within bounds, return df.columns[index_fallback].
      - Else return None.
    """
    if preferred_name in df.columns:
        return preferred_name
    if index_fallback is not None and 0 <= index_fallback < len(df.columns):
        return df.columns[int(index_fallback)]
    return None

# ---------- UI ----------
uploaded = st.file_uploader("Upload password-protected .xlsx", type=["xlsx"])
password = st.text_input("Password", type="password", value="_S8&Dwy2&U")

with st.expander("Advanced (column fallbacks)"):
    st.markdown("If your columns are unnamed, the app will fall back to these positions.")
    bw_idx = st.number_input("Fallback index for BW (manufacture_name)", min_value=0, value=74, step=1)
    bx_idx = st.number_input("Fallback index for BX (manufacture_address)", min_value=0, value=75, step=1)
    bz_idx = st.number_input("Fallback index for BZ (optional extra truncation)", min_value=0, value=77, step=1)
    coo_idx = st.number_input("Fallback index for country_of_origin (Unnamed: 63)", min_value=0, value=63, step=1)
    unnamed_state_idx = st.number_input("Fallback index for STATE unnamed column (commonly Unnamed: 23)", min_value=0, value=23, step=1)

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

        # Determine columns (prefer named; fallback to "Unnamed:X"; then index)
        bw_col = safe_get_col_by_name_or_index(df_hawb, "manufacture_name", bw_idx)
        if bw_col is None:
            bw_col = safe_get_col_by_name_or_index(df_hawb, "Unnamed: 74", bw_idx)

        bx_col = safe_get_col_by_name_or_index(df_hawb, "manufacture_address", bx_idx)
        if bx_col is None:
            bx_col = safe_get_col_by_name_or_index(df_hawb, "Unnamed: 75", bx_idx)

        bz_col = safe_get_col_by_name_or_index(df_hawb, "Unnamed: 77", bz_idx)

        # Apply truncation rules
        if bw_col in df_hawb.columns:
            df_hawb[bw_col] = df_hawb[bw_col].apply(lambda x: truncate_half_if_over(x, 100))
        if bx_col in df_hawb.columns:
            df_hawb[bx_col] = df_hawb[bx_col].apply(lambda x: truncate_half_if_over(x, 225))
        if bz_col in df_hawb.columns:
            df_hawb[bz_col] = df_hawb[bz_col].apply(lambda x: truncate_half_if_over(x, 8))

        # ---- Set ALL country_of_origin to "CN" ----
        coo_col = get_col_fuzzy(df_hawb, ["country_of_origin", "unnamed: 63"], coo_idx)
        if coo_col in df_hawb.columns:
            df_hawb[coo_col] = "CN"

        # ---- Remove STATE or unnamed column at index 23 ----
        to_drop = []
        if "STATE" in df_hawb.columns:
            to_drop.append("STATE")
        if 0 <= unnamed_state_idx < len(df_hawb.columns):
            fallback_colname = df_hawb.columns[int(unnamed_state_idx)]
            if (
                fallback_colname in df_hawb.columns
                and fallback_colname not in to_drop
                and (str(fallback_colname).startswith("Unnamed:") or fallback_colname == "Unnamed: 23")
            ):
                to_drop.append(fallback_colname)
        if to_drop:
            df_hawb.drop(columns=to_drop, inplace=True, errors="ignore")

        # ---------- MAWB ----------
        if "mawb" not in xls.sheet_names:
            st.error("Sheet 'mawb' not found in the workbook.")
            st.stop()
        df_mawb = pd.read_excel(xls, sheet_name="mawb")

        # Set L2 (Excel) => row index 0 in pandas for column 'consignee_id_number'
        if "consignee_id_number" not in df_mawb.columns:
            # Create/ensure at least one row, then set value
            if len(df_mawb) == 0:
                df_mawb = pd.DataFrame({"consignee_id_number": ["2567704"]})
            else:
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

        # Optional previews
        with st.expander("Preview (first 10 rows) – hawb"):
            st.dataframe(df_hawb.head(10))
        with st.expander("Preview (first 10 rows) – mawb"):
            st.dataframe(df_mawb.head(10))

    except Exception as e:
        st.error(f"Processing failed: {e}")
else:
    st.info("Upload a file, enter the password, then click **Process**.")
