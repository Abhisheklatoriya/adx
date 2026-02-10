import streamlit as st
import pandas as pd
import re
import zipfile
import io

# 1. Expand limits and set layout
st.set_page_config(page_title="Ad creative matcher", layout="wide")
st.title("⚡ Ultra-Fast Ad Matcher (XLSX)")

AD_CODE_RE = re.compile(r"\b\d{8}\b")


def _digits_only(x: str) -> str:
    return re.sub(r"\D", "", str(x or ""))


# 2. Cache the Excel extraction (auto-detect header row + find Ad Codes)
@st.cache_data
def get_ad_data_xlsx(uploaded_file):
    """
    Returns:
      - codes: sorted unique list of 8-digit ad codes (as strings)
      - df: the parsed dataframe with normalized column names
      - ad_code_col: the detected Ad Code column name (normalized)
    """
    # Read as bytes so caching works reliably
    content = uploaded_file.getvalue()
    bio = io.BytesIO(content)

    # 1) Read without headers to detect the real header row (your file has title rows)
    raw = pd.read_excel(bio, engine="openpyxl", header=None)

    # Find a row that contains something like "Ad Code"
    header_row_idx = None
    for i in range(min(len(raw), 80)):  # scan top 80 rows
        row = raw.iloc[i].astype(str).str.strip().str.lower()
        if row.str.contains(r"\bad\s*code\b", regex=True, na=False).any():
            header_row_idx = i
            break

    if header_row_idx is None:
        # fallback: assume first row is header
        header_row_idx = 0

    # 2) Re-read using that row as header
    bio2 = io.BytesIO(content)
    df = pd.read_excel(bio2, engine="openpyxl", header=header_row_idx)

    # Normalize column names
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Try to find "ad code" column robustly
    ad_code_col = None
    for c in df.columns:
        if re.search(r"\bad\s*code\b", c):
            ad_code_col = c
            break

    if ad_code_col is None:
        # fallback: scan all cells for 8-digit codes
        all_text = "\n".join(
            df.astype(str).fillna("").values.ravel().tolist()
        )
        codes = sorted(set(AD_CODE_RE.findall(all_text)))
        return codes, df, None

    # Extract codes from the ad code column, normalize to 8-digit strings
    col = df[ad_code_col].astype(str).fillna("")
    extracted = col.apply(lambda x: _digits_only(x)[-8:])  # last 8 digits (safe)
    codes = sorted(set([c for c in extracted.tolist() if len(c) == 8]))

    return codes, df, ad_code_col


# 3. Cache the Assets in memory so they don't reload on every click
@st.cache_resource
def load_assets(uploaded_files):
    processed_assets = []
    for uploaded_file in uploaded_files:
        if uploaded_file.name.lower().endswith(".zip"):
            with zipfile.ZipFile(uploaded_file) as z:
                for file_info in z.infolist():
                    if file_info.is_dir() or "__MACOSX" in file_info.filename:
                        continue
                    with z.open(file_info) as f:
                        data = f.read()
                        processed_assets.append(
                            {
                                "name": file_info.filename,
                                "data": data,
                                "ext": file_info.filename.split(".")[-1].lower(),
                            }
                        )
        else:
            processed_assets.append(
                {
                    "name": uploaded_file.name,
                    "data": uploaded_file.getvalue(),
                    "ext": uploaded_file.name.split(".")[-1].lower(),
                }
            )
    return processed_assets


def get_rows_for_code(df: pd.DataFrame, ad_code_col: str | None, code: str) -> pd.DataFrame:
    """
    Returns matching rows for an 8-digit code, using robust normalization.
    If ad_code_col is missing, returns empty df.
    """
    if df is None or ad_code_col is None or ad_code_col not in df.columns:
        return df.iloc[0:0] if df is not None else pd.DataFrame()

    code_digits = _digits_only(code)[-8:]

    col_digits = (
        df[ad_code_col]
        .astype(str)
        .fillna("")
        .apply(lambda x: _digits_only(x)[-8:])
    )

    return df[col_digits == code_digits]


def preview_asset(asset):
    ext = asset["ext"]
    if ext in ["mp4", "mov", "webm"]:
        st.video(asset["data"])
    elif ext in ["mp3", "wav"]:
        st.audio(asset["data"])
    elif ext in ["jpg", "jpeg", "png", "gif", "webp"]:
        st.image(asset["data"])
    else:
        st.info(f"No inline preview for .{ext}. You can still download it.")


# Sidebar Uploads
with st.sidebar:
    st.header("Upload")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    raw_files = st.file_uploader("Upload Assets or ZIP", accept_multiple_files=True)
    if st.button("Clear Cache / Reset"):
        st.cache_resource.clear()
        st.cache_data.clear()
        st.rerun()

if excel_file and raw_files:
    # Load once
    all_assets = load_assets(raw_files)
    ad_codes, df, ad_code_col = get_ad_data_xlsx(excel_file)

    st.success(f"Loaded {len(all_assets)} files. Found {len(ad_codes)} Ad Codes.")

    # Fast Search
    search = st.text_input("🔍 Quick Search Ad Code", placeholder="Type to filter...")
    filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes

    # Helpful debug if nothing shows
    with st.expander("🛠 Debug (open if nothing is matching)", expanded=False):
        st.write("Detected Ad Code column:", ad_code_col)
        st.write("First 20 extracted codes:", filtered_codes[:20])
        st.write("First 20 asset names:", [a["name"] for a in all_assets][:20])

    for code in filtered_codes:
        code_digits = _digits_only(code)[-8:]

        # Robust matching: compare digits-only filename vs digits-only code
        matches = []
        for a in all_assets:
            fname_digits = _digits_only(a["name"])
            if code_digits and code_digits in fname_digits:
                matches.append(a)

        if matches:
            with st.expander(f"✅ Ad {code_digits} - {len(matches)} files found", expanded=True):
                c1, c2 = st.columns([1, 1.5])

                with c1:
                    st.markdown("**Ad Details (from Excel):**")

                    rows = get_rows_for_code(df, ad_code_col, code_digits)
                    if rows is not None and len(rows) > 0:
                        # Show the matching row(s)
                        st.dataframe(rows, use_container_width=True, hide_index=True)

                        # Also show as key-value for quick copy
                        st.markdown("**Key fields (first match):**")
                        first = rows.iloc[0].to_dict()
                        # Remove NaNs / blanks for readability
                        cleaned = {k: v for k, v in first.items() if str(v).strip() not in ["nan", "None", ""]}
                        st.json(cleaned)
                    else:
                        st.warning("No matching rows found in Excel for this Ad Code (check formatting).")

                with c2:
                    st.markdown("**Creative Previews + Downloads:**")
                    for asset in matches:
                        st.caption(f"File: {asset['name']}")
                        preview_asset(asset)
                        st.download_button(
                            "Download",
                            data=asset["data"],
                            file_name=asset["name"],
                            key=f"dl_{asset['name']}_{code_digits}",
                        )
                        st.divider()
        # If no matches, we skip rendering (keeps UI clean)

else:
    st.info("Please upload your Excel file and creative assets to begin.")
