import io
import re
from typing import Optional

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
   - If `Unnamed: 77` (≈ **BZ**) length > 8, keep first half.
   - Fill missing `country_of_origin` (or `Unnamed: 63`) with **"CN"**.
   - **If `manufacture_zip_code` (or `Unnamed: 78`) is not exactly 6 digits, set to "123456".**
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
    decrypted = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(uploaded_file)
    office_file.load_key(password=password)
    office_file.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted

def normalize_name(name: str) -> str:
    """lowercase, strip, remove non-alphanum to make matching resilient to spaces/underscores."""
    s = str(name).strip().lower()
    return re.sub(r"[^a-z0-9]+", "", s)

def safe_get_col_by_name_or_index(df: pd.DataFrame, preferred_name: str, index_fallback: Optional[int]):
    """Try exact name (after strip), then index fallback."""
    cols = [str(c).strip() for c in df.columns]
    if preferred_name.strip() in cols:
        return preferred_name.strip()
    if index_fallback is not None and 0 <= int(index_fallback) < len(cols):
        return cols[int(index_fallback)]
    return None

def get_col_fuzzy(df: pd.DataFrame, preferred_names: list[str], index_fallback: Optional[int] = None):
    """
    Try exact (trimmed) names, then fuzzy normalized names, then explicit 'Unnamed: N',
    finally raw index fallback.
    """
    # Trim headers once
    df.columns = pd.Index([str(c).strip() for c in df.columns])

    # 1) exact
    for name in preferred_names:
        if name.strip() in df.columns:
            return name.strip()

    # 2) fuzzy normalized match
    targets = [normalize_name(n) for n in preferred_names]
    for col in df.columns:
        if normalize_name(col) in targets:
            return col

    # 3) handle "Unnamed: N" literal in preferred_names
    for name in preferred_names:
        m = re.fullmatch(r"unnamed:\s*(\d+)", name.strip().lower())
        if m:
            n = int(m.group(1))
            # exact literal header match
            literal = f"Unnamed: {n}"
            if literal in df.columns:
                return literal
            # or column at 0-based index n if present
            if 0 <= n < len(df.columns):
                return df.columns[n]

    # 4) raw index fallback
    if index_fallback is not None and 0 <= int(index_fallback) < len(df.columns):
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
    zip_idx = st.number_input("Fallback index for manufacture_zip_code (Unnamed: 78)", min_value=0, value=78, step=1)
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
        # Trim headers to avoid trailing spaces issues
        df_hawb.columns = pd.Index([str(c).strip() for c in df_hawb.columns])

        # Determine columns (prefer named; fallback)
        bw_col = get_col_fuzzy(df_hawb, ["manufacture_name", "unnamed: 74"], bw_idx)
        bx_col = get_col_fuzzy(df_hawb, ["manufacture_address", "unnamed: 75"], bx_idx)
        bz_col = get_col_fuzzy(df_hawb, ["unnamed: 77"], bz_idx)

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

        # ---- Zip code rule: set "123456" if not exactly 6 digits ----
        zip_col = get_col_fuzzy(df_hawb, ["manufacture_zip_code", "manufacture_zip_code ", "unnamed: 78"], zip_idx)
        before_zip_sample = None
        if zip_col in df_hawb.columns:
            # sample before
            before_zip_sample = df_hawb[zip_col].head(5).tolist()

            # normalize dtype to string for regex
            df_hawb[zip_col] = df_hawb[zip_col].astype("string")

            # Treat blanks as NA for easier logic
            df_hawb[zip_col] = df_hawb[zip_col].replace(r"^\s*$", pd.NA, regex=True)

            # Valid = exactly 6 digits
            is_valid = df_hawb[zip_col].str.match(r"^\d{6}$", na=False)

            # Everything else -> "123456"
            df_hawb.loc[~is_valid, zip_col] = "123456"

        # ---- Remove STATE or unnamed column at index 23 ----
        to_drop = []
        if "STATE" in df_hawb.columns:
            to_drop.append("STATE")
        # unnamed fallback (avoid dropping zip column by accident)
        if 0 <= int(unnamed_state_idx) < len(df_hawb.columns):
            fallback_colname = df_hawb.columns[int(unnamed_state_idx)]
            if (
                fallback_colname in df_hawb.columns
                and fallback_colname not in to_drop
                and (str(fallback_colname).startswith("Unnamed:") or fallback_colname == "Unnamed: 23")
                and fallback_colname != zip_col
            ):
                to_drop.append(fallback_colname)
        if to_drop:
            df_hawb.drop(columns=to_drop, inplace=True, errors="ignore")

        # ---------- MAWB ----------
        if "mawb" not in xls.sheet_names:
            st.error("Sheet 'mawb' not found in the workbook.")
            st.stop()
        df_mawb = pd.read_excel(xls, sheet_name="mawb")
        df_mawb.columns = pd.Index([str(c).strip() for c in df_mawb.columns])

        # Set L2 (Excel) => row index 0 in pandas for column 'consignee_id_number'
        if "consignee_id_number" not in df_mawb.columns:
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

        # Debug + previews
        with st.expander("Debug / Column selections"):
            st.write({
                "bw_col": bw_col,
                "bx_col": bx_col,
                "bz_col": bz_col,
                "coo_col": coo_col,
                "zip_col": zip_col,
                "dropped_columns": to_drop,
            })
            if zip_col in df_hawb.columns:
                st.write("Zip sample before:", before_zip_sample)
                st.write("Zip sample after:", df_hawb[zip_col].head(5).tolist())

        with st.expander("Preview (first 10 rows) – hawb"):
            st.dataframe(df_hawb.head(10))
        with st.expander("Preview (first 10 rows) – mawb"):
            st.dataframe(df_mawb.head(10))

    except Exception as e:
        st.error(f"Processing failed: {e}")
else:
    st.info("Upload a file, enter the password, then click **Process**.")
