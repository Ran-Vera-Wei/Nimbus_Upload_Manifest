import io
import re
from typing import Optional, List

import numpy as np
import pandas as pd
import msoffcrypto
import streamlit as st

st.set_page_config(page_title="Excel Converter", page_icon="📄")
st.title("📄 Excel Converter (Password Remove + Transform)")

st.markdown("""
**Rules (apply to `hawb`, starting from row 3):**
- `manufacture_name` : if length > 100 → keep first half
- `manufacture_address` : if length > 225 → keep first half
- `manufacture_state`: if length > 8 → keep first half
- **Set all `country_of_origin` to "CN"**
- **Set all `manufacture_country` to "CN"**
- **Zip** (`manufacture_zip_code` or `Unnamed: 78`): if not exactly 6 digits → **"123456"**
- Drop `STATE` column (or unnamed index 23) entirely

**Rules (apply to `Mawb`):**
- set L2 (`consignee_id_number`) to `2567704`
""")

# ---------- Helpers ----------
def decrypt_xlsx(uploaded_file, password: str) -> io.BytesIO:
    buf = io.BytesIO()
    of = msoffcrypto.OfficeFile(uploaded_file)
    of.load_key(password=password)
    of.decrypt(buf)
    buf.seek(0)
    return buf

def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())

def find_col_index_from_header_row(header_row: pd.Series, candidates: List[str], index_fallback: Optional[int]=None) -> Optional[int]:
    """
    Given the SECOND row (real headers), find the column index matching:
      exact -> normalized -> 'Unnamed: N' literal or that position -> fallback index
    Returns 0-based column index or None.
    """
    headers = [str(x).strip() for x in header_row.values]
    # exact
    for c in candidates:
        if str(c).strip() in headers:
            return headers.index(str(c).strip())
    # normalized fuzzy
    target_norms = {norm(c) for c in candidates}
    for j, h in enumerate(headers):
        if norm(h) in target_norms:
            return j
    # explicit Unnamed:N
    for c in candidates:
        m = re.fullmatch(r"unnamed:\s*(\d+)", str(c).strip().lower())
        if m:
            n = int(m.group(1))
            literal = f"Unnamed: {n}"
            if literal in headers:
                return headers.index(literal)
            if 0 <= n < len(headers):
                return n
    # fallback index
    if index_fallback is not None and 0 <= int(index_fallback) < len(headers):
        return int(index_fallback)
    return None

def truncate_half_if_over_val(x, thresh: int):
    if isinstance(x, str) and len(x) > thresh:
        return x[: len(x)//2]
    return x

# ---------- UI ----------
uploaded = st.file_uploader("Upload password-protected .xlsx", type=["xlsx"])
password = st.text_input("Password", type="password", value="_S8&Dwy2&U")

with st.expander("Advanced (0-based index fallbacks)"):
    bw_idx = st.number_input("BW (manufacture_name) fallback index", min_value=0, value=74, step=1)
    bx_idx = st.number_input("BX (manufacture_address) fallback index", min_value=0, value=75, step=1)
    bz_idx = st.number_input("BZ (optional extra truncation) fallback index", min_value=0, value=77, step=1)
    coo_idx = st.number_input("country_of_origin fallback index (≈ Unnamed: 63)", min_value=0, value=63, step=1)
    mao_idx = st.number_input("country_of_origin fallback index (≈ Unnamed: 79)", min_value=0, value=79, step=1)
    zip_idx = st.number_input("manufacture_zip_code fallback index (≈ Unnamed: 78)", min_value=0, value=78, step=1)
    unnamed_state_idx = st.number_input("Unnamed state col index (commonly 23)", min_value=0, value=23, step=1)

if st.button("Process") and uploaded and password:
    try:
        dec = decrypt_xlsx(uploaded, password)
        xls = pd.ExcelFile(dec, engine="openpyxl")

        # ---------- HAWB: keep both header rows intact ----------
        if "hawb" not in xls.sheet_names:
            st.error("Sheet 'hawb' not found.")
            st.stop()

        # Read raw to keep rows 1&2 untouched
        hawb_raw = pd.read_excel(xls, sheet_name="hawb", header=None, dtype=object)
        if hawb_raw.shape[0] < 2:
            st.error("`hawb` has fewer than 2 rows.")
            st.stop()

        header2 = hawb_raw.iloc[1]  # second row = real headers
        data_start = 2              # we edit from row index 2 (third row)

        # find needed column indices by header2
        bw_j = find_col_index_from_header_row(header2, ["manufacture_name", "unnamed: 74"], bw_idx)
        bx_j = find_col_index_from_header_row(header2, ["manufacture_address", "unnamed: 75"], bx_idx)
        bz_j = find_col_index_from_header_row(header2, ["manufacture_state", "unnamed: 77"], bz_idx)
        coo_j = find_col_index_from_header_row(header2, ["country_of_origin", "unnamed: 63"], coo_idx)
        mao_j = find_col_index_from_header_row(header2, ["manufacture_country", "unnamed: 79"], mao_idx)
        zip_j = find_col_index_from_header_row(header2, ["manufacture_zip_code", "manufacture_zip_code ", "unnamed: 78"], zip_idx)

        # Apply rules directly on hawb_raw rows >= data_start
        # Truncations
        if bw_j is not None and bw_j < hawb_raw.shape[1]:
            hawb_raw.iloc[data_start:, bw_j] = hawb_raw.iloc[data_start:, bw_j].apply(lambda v: truncate_half_if_over_val(v, 100))
        if bx_j is not None and bx_j < hawb_raw.shape[1]:
            hawb_raw.iloc[data_start:, bx_j] = hawb_raw.iloc[data_start:, bx_j].apply(lambda v: truncate_half_if_over_val(v, 225))
        if bz_j is not None and bz_j < hawb_raw.shape[1]:
            hawb_raw.iloc[data_start:, bz_j] = hawb_raw.iloc[data_start:, bz_j].apply(lambda v: truncate_half_if_over_val(v, 8))

        # country_of_origin -> "CN" for all rows from row 3
        if coo_j is not None and coo_j < hawb_raw.shape[1]:
            hawb_raw.iloc[data_start:, coo_j] = "CN"
        
        # manufacture_country -> "CN" for all rows from row 3
        if mao_j is not None and mao_j < hawb_raw.shape[1]:
            hawb_raw.iloc[data_start:, mao_j] = "CN"

        # zip: non-6-digits -> "123456" (rows from row 3)
        if zip_j is not None and zip_j < hawb_raw.shape[1]:
            col = hawb_raw.iloc[data_start:, zip_j].astype("string")
            col = col.replace(r"^\s*$", pd.NA, regex=True)
            valid = col.str.match(r"^\d{6}$", na=False)
            col.loc[~valid] = "123456"
            hawb_raw.iloc[data_start:, zip_j] = col

        # Drop STATE column (by checking header2) OR unnamed index 23
        drop_indices = []
        for j, name in enumerate([str(x).strip() for x in header2.values]):
            if name == "STATE":
                drop_indices.append(j)
        if 0 <= int(unnamed_state_idx) < hawb_raw.shape[1] and int(unnamed_state_idx) not in drop_indices:
            drop_indices.append(int(unnamed_state_idx))
        if drop_indices:
            keep = [j for j in range(hawb_raw.shape[1]) if j not in drop_indices]
            hawb_raw = hawb_raw.iloc[:, keep]

        # ---------- MAWB ----------
        if "mawb" not in xls.sheet_names:
            st.error("Sheet 'mawb' not found.")
            st.stop()

        df_mawb = pd.read_excel(xls, sheet_name="mawb")
        df_mawb.columns = pd.Index([str(c).strip() for c in df_mawb.columns])

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

        # ---------- Write workbook ----------
        out_buf = io.BytesIO()
        with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
            # Write hawb exactly as matrix (keeps both header rows)
            hawb_raw.to_excel(writer, sheet_name="hawb", header=False, index=False)
            df_mawb.to_excel(writer, sheet_name="mawb", index=False)
        out_buf.seek(0)

        st.success("Done! Download your converted file below.")
        st.download_button(
            "⬇️ Download converted.xlsx",
            data=out_buf,
            file_name="converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Debug"):
            st.write({
                "hawb_shape": tuple(hawb_raw.shape),
                "bw_idx": bw_j, "bx_idx": bx_j, "bz_idx": bz_j,
                "coo_idx": coo_j, "zip_idx": zip_j,
                "dropped_indices": drop_indices
            })
            st.write("Row 2 headers (real):", [str(x).strip() for x in header2.values])

    except Exception as e:
        st.error(f"Processing failed: {e}")
else:
    st.info("Upload a file, enter the password, then click **Process**.")
