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
2. In **`hawb`** (sheet with **2 header rows**):
   - Keeps the **first two rows unchanged**.
   - From the **3rd row onward**:
     - If `manufacture_name` (**BW**) length > 100 → keep first half.
     - If `manufacture_address` (**BX**) length > 225 → keep first half.
     - If **BZ** (often `Unnamed: 77`) length > 8 → keep first half. *(optional rule)*
     - **Set ALL `country_of_origin` to "CN".**
     - If `manufacture_zip_code` (or fallback) is **not exactly 6 digits**, set to **"123456"**.
   - Remove the `STATE` column (entire column) or unnamed column at index **23** (0-based) across the whole sheet.
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

def norm(s: str) -> str:
    """Lowercase+strip and remove non-alphanum to match headers robustly."""
    s = str(s).strip().lower()
    return re.sub(r"[^a-z0-9]+", "", s)

def get_col_by_header_row(header_row: pd.Series, candidates: list[str], index_fallback: Optional[int] = None) -> Optional[str]:
    """
    Find a column name using the provided header row values (row 2 in Excel),
    trying exact match → fuzzy normalized match → 'Unnamed: N' → index fallback.
    Returns the picked column name (as in header_row), or None.
    """
    # ensure header row elements are strings and trimmed
    headers = pd.Index([str(x).strip() for x in header_row.values])

    # 1) exact
    for c in candidates:
        c_trim = str(c).strip()
        if c_trim in headers:
            return c_trim

    # 2) fuzzy normalized
    targets = {norm(c) for c in candidates}
    for h in headers:
        if norm(h) in targets:
            return h

    # 3) explicit "Unnamed: N" (by text or by position N)
    for c in candidates:
        m = re.fullmatch(r"unnamed:\s*(\d+)", str(c).strip().lower())
        if m:
            n = int(m.group(1))
            literal = f"Unnamed: {n}"
            if literal in headers:
                return literal
            if 0 <= n < len(headers):
                return headers[n]

    # 4) raw index fallback
    if index_fallback is not None and 0 <= int(index_fallback) < len(headers):
        return headers[int(index_fallback)]

    return None

# ---------- UI ----------
uploaded = st.file_uploader("Upload password-protected .xlsx", type=["xlsx"])
password = st.text_input("Password", type="password", value="_S8&Dwy2&U")

with st.expander("Advanced (column fallbacks)"):
    st.markdown("If your columns are unnamed or shifting, the app will fall back to these positions (0-based).")
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

        # ---------- HAWB (two header rows) ----------
        if "hawb" not in xls.sheet_names:
            st.error("Sheet 'hawb' not found in the workbook.")
            st.stop()

        # Read WITHOUT header to preserve the first two rows exactly
        hawb_raw = pd.read_excel(xls, sheet_name="hawb", header=None, dtype=object)

        if hawb_raw.shape[0] < 3:
            st.error("The 'hawb' sheet seems to have fewer than 3 rows; cannot apply row-3-and-below rules.")
            st.stop()

        # Row 0 = Unnamed..., Row 1 = real headers, Row 2+ = data
        header_row = hawb_raw.iloc[1]  # second row as real headers
        data = hawb_raw.iloc[2:].copy()
        # Use row-1 as the column headers for data operations
        data.columns = [str(x).strip() for x in header_row.values]

        # Determine columns from header_row
        bw_col = get_col_by_header_row(header_row, ["manufacture_name", "unnamed: 74"], bw_idx)
        bx_col = get_col_by_header_row(header_row, ["manufacture_address", "unnamed: 75"], bx_idx)
        bz_col = get_col_by_header_row(header_row, ["unnamed: 77"], bz_idx)
        coo_col = get_col_by_header_row(header_row, ["country_of_origin", "unnamed: 63"], coo_idx)
        zip_col = get_col_by_header_row(header_row, ["manufacture_zip_code", "manufacture_zip_code ", "unnamed: 78"], zip_idx)

        # Apply truncations on DATA ONLY (from 3rd row)
        if bw_col in data.columns:
            data[bw_col] = data[bw_col].apply(lambda x: truncate_half_if_over(x, 100))
        if bx_col in data.columns:
            data[bx_col] = data[bx_col].apply(lambda x: truncate_half_if_over(x, 225))
        if bz_col in data.columns:
            data[bz_col] = data[bz_col].apply(lambda x: truncate_half_if_over(x, 8))

        # Set ALL country_of_origin to "CN" (from 3rd row)
        if coo_col in data.columns:
            data[coo_col] = "CN"

        # Zip rule: set "123456" if not exactly 6 digits (from 3rd row)
        if zip_col in data.columns:
            data[zip_col] = data[zip_col].astype("string")
            data[zip_col] = data[zip_col].replace(r"^\s*$", pd.NA, regex=True)
            valid = data[zip_col].str.match(r"^\d{6}$", na=False)
            data.loc[~valid, zip_col] = "123456"

        # Reassemble the hawb sheet with first two rows preserved
        # Make sure data columns are ordered like header_row (to align with original structure)
        ordered_cols = [str(x).strip() for x in header_row.values]
        data_ordered = data.reindex(columns=ordered_cols, copy=False)

        hawb_out = pd.concat(
            [hawb_raw.iloc[:2].copy(), data_ordered.reset_index(drop=True)],
            ignore_index=True
        )

        # Remove STATE column OR unnamed 23 across entire output (drop the whole column)
        # We need to identify the column index to drop. We'll use header_row to find the position.
        drop_col_indices = []
        # Drop where header row equals "STATE"
        for j, name in enumerate(ordered_cols):
            if str(name).strip() == "STATE":
                drop_col_indices.append(j)
        # Also allow unnamed-state fallback at index unnamed_state_idx
        if 0 <= int(unnamed_state_idx) < len(ordered_cols):
            if int(unnamed_state_idx) not in drop_col_indices:
                drop_col_indices.append(int(unnamed_state_idx))

        if drop_col_indices:
            keep_indices = [j for j in range(hawb_out.shape[1]) if j not in drop_col_indices]
            hawb_out = hawb_out.iloc[:, keep_indices]

        # ---------- MAWB (assumed single header row) ----------
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
            # hawb_out already includes 2 header rows; write without pandas header
            hawb_out.to_excel(writer, sheet_name="hawb", header=False, index=False)
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
                "dropped_col_indices": drop_col_indices,
                "hawb_shape_before": tuple(hawb_raw.shape),
                "hawb_shape_after": tuple(hawb_out.shape),
            })

        with st.expander("Preview (first 10 data rows) – hawb (from row 3)"):
            st.dataframe(data_ordered.head(10))
        with st.expander("Preview (first 10 rows) – mawb"):
            st.dataframe(df_mawb.head(10))

    except Exception as e:
        st.error(f"Processing failed: {e}")
else:
    st.info("Upload a file, enter the password, then click **Process**.")
