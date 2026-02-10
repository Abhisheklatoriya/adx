import streamlit as st
import pandas as pd
import re
import zipfile
import io

# 1) Expand limits and set layout
st.set_page_config(page_title="Ad creative matcher", layout="wide")
st.title("⚡ Ultra-Fast Ad Matcher (XLSX)")

AD_CODE_RE = re.compile(r"\b\d{8}\b")


# 2) Cache the Excel extraction
@st.cache_data
def get_ad_data(file):
    """
    Reads Excel and returns:
      - ad_codes: sorted unique list of 8-digit codes found in the sheet
      - df: cleaned dataframe (with original columns)
      - ad_code_col: detected column name that contains 'Ad Code' (best effort)
    """
    # IMPORTANT: requires openpyxl in requirements.txt on Streamlit Cloud
    df = pd.read_excel(file, engine="openpyxl")

    # Normalize columns for detection (keep original df columns as-is)
    cols_lower = {str(c).strip().lower(): c for c in df.columns}

    # Best-effort: find a column that looks like "Ad Code"
    ad_code_col = None
    for lower, original in cols_lower.items():
        if "ad" in lower and "code" in lower:
            ad_code_col = original
            break

    # Extract ad codes:
    # If ad_code_col exists, prefer it (clean to 8 digits).
    # Else, search the whole sheet text for 8-digit codes.
    codes = set()

    if ad_code_col is not None:
        series = df[ad_code_col].astype(str)
        for v in series:
            m = AD_CODE_RE.search(v)
            if m:
                codes.add(m.group(0))
    else:
        # fallback: scan all cells
        flat_text = "\n".join(df.astype(str).fillna("").values.flatten().tolist())
        codes.update(AD_CODE_RE.findall(flat_text))

    ad_codes = sorted(codes)
    return ad_codes, df, ad_code_col


# 3) Cache the Assets in memory so they don't reload on every click
@st.cache_resource
def load_assets(uploaded_files):
    processed_assets = []
    for uploaded_file in uploaded_files:
        if uploaded_file.name.lower().endswith(".zip"):
            # Read zip into memory (works for small/medium zips)
            zip_bytes = uploaded_file.getvalue()
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
                for file_info in z.infolist():
                    if file_info.is_dir() or "__MACOSX" in file_info.filename:
                        continue
                    with z.open(file_info) as f:
                        data = f.read()
                    processed_assets.append(
                        {
                            "name": file_info.filename,
                            "data": data,
                            "ext": file_info.filename.split(".")[-1].lower()
                            if "." in file_info.filename
                            else "",
                        }
                    )
        else:
            processed_assets.append(
                {
                    "name": uploaded_file.name,
                    "data": uploaded_file.getvalue(),
                    "ext": uploaded_file.name.split(".")[-1].lower()
                    if "." in uploaded_file.name
                    else "",
                }
            )
    return processed_assets


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
    all_assets = load_assets(raw_files)
    ad_codes, df, ad_code_col = get_ad_data(excel_file)

    st.success(f"Loaded {len(all_assets)} files. Found {len(ad_codes)} Ad Codes.")

    # Fast Search
    search = st.text_input("🔍 Quick Search Ad Code", placeholder="Type to filter...")
    filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes

    for code in filtered_codes:
        # Match based on filename containing the code
        matches = [a for a in all_assets if code in a["name"]]

        if matches:
            with st.expander(f"✅ Ad {code} - {len(matches)} files found", expanded=True):
                c1, c2 = st.columns([1, 1.5])

                # LEFT: Ad Specs from Excel
                with c1:
                    st.markdown("**Ad Specs:**")

                    if ad_code_col is not None:
                        # try exact match in column (as text)
                        col_as_text = df[ad_code_col].astype(str).str.extract(r"(\d{8})", expand=False)
                        rows = df[col_as_text == code]
                        if len(rows) > 0:
                            st.dataframe(rows, use_container_width=True)
                        else:
                            # fallback: show any rows where code appears anywhere
                            mask = df.astype(str).apply(lambda r: r.str.contains(code, regex=False, na=False)).any(axis=1)
                            rows2 = df[mask]
                            st.dataframe(rows2 if len(rows2) else df.head(5), use_container_width=True)
                    else:
                        # No ad code col detected, show rows containing the code anywhere
                        mask = df.astype(str).apply(lambda r: r.str.contains(code, regex=False, na=False)).any(axis=1)
                        rows = df[mask]
                        st.dataframe(rows if len(rows) else df.head(5), use_container_width=True)

                # RIGHT: Creative previews + downloads
                with c2:
                    for asset in matches:
                        st.caption(f"File: {asset['name']}")

                        if asset["ext"] in ["mp4", "mov", "webm"]:
                            st.video(asset["data"])
                        elif asset["ext"] in ["mp3", "wav"]:
                            st.audio(asset["data"])
                        elif asset["ext"] in ["jpg", "jpeg", "png", "gif", "webp"]:
                            st.image(asset["data"])

                        st.download_button(
                            "Download",
                            data=asset["data"],
                            file_name=asset["name"].split("/")[-1],
                            key=f"dl_{asset['name']}_{code}",
                        )
                        st.divider()

else:
    st.info("Please upload your Excel file and creative assets to begin.")
