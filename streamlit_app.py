import streamlit as st
import pandas as pd
import re
import zipfile
import io
from typing import Tuple, List, Optional

# 1) Page config
st.set_page_config(page_title="Ad creative matcher", layout="wide")
st.title("⚡ Ultra-Fast Ad Matcher (XLSX)")

CODE_RE = re.compile(r"\b(\d{8})\b")


# -----------------------------
# XLSX parsing (replaces DOCX)
# - Auto-detects the header row that contains "Ad Code"
# - Returns ad_codes + dataframe + ad_code_col
# -----------------------------
@st.cache_data
def get_ad_data_xlsx(file) -> Tuple[List[str], pd.DataFrame, str]:
    # Read raw to find header row (your sheet has title rows above the real header)
    raw = pd.read_excel(file, header=None, engine="openpyxl")

    header_row = None
    for i in range(min(80, len(raw))):
        row = raw.iloc[i].astype(str).str.strip().str.lower()
        if row.str.contains(r"\bad\s*code\b", regex=True).any():
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not find a header row containing 'Ad Code' in the first 80 rows.")

    df = pd.read_excel(file, header=header_row, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    # Find the ad code column
    ad_code_col = None
    for c in df.columns:
        cl = c.lower().strip()
        if "ad" in cl and "code" in cl:
            ad_code_col = c
            break

    if not ad_code_col:
        raise ValueError(f"Header found but no Ad Code column. Columns: {list(df.columns)}")

    # Normalize ad codes
    df[ad_code_col] = df[ad_code_col].astype(str).str.strip()
    df = df[df[ad_code_col].str.match(r"^\d{8}$")]

    ad_codes = sorted(df[ad_code_col].unique().tolist())
    return ad_codes, df, ad_code_col


# -----------------------------
# Asset loading
# - Original behavior: load everything into memory
# - Safe mode: do NOT keep bytes for big ZIPs; store only file handles (zip entry pointers)
#   (Still allows preview + download, but avoids RAM blowups)
# -----------------------------
@st.cache_resource
def load_assets(uploaded_files, safe_mode: bool):
    """
    Returns a list of asset dicts.
    In safe_mode:
      - For ZIP entries: store (zip_bytes, entry_name) instead of fully reading bytes.
      - For normal files: store bytes (usually smaller), but you can extend to stream too.
    """
    processed_assets = []

    for uploaded_file in uploaded_files:
        name_lower = uploaded_file.name.lower()

        if name_lower.endswith(".zip"):
            zip_bytes = uploaded_file.getvalue()  # NOTE: still reads ZIP into memory once
            # If your zips are huge, you should not use this approach (disk-based is needed).
            zf = zipfile.ZipFile(io.BytesIO(zip_bytes))

            for info in zf.infolist():
                if info.is_dir() or "__MACOSX" in info.filename:
                    continue

                ext = info.filename.split(".")[-1].lower() if "." in info.filename else ""

                if safe_mode:
                    # Don't read file bytes yet; store reference
                    processed_assets.append({
                        "name": info.filename,
                        "ext": ext,
                        "kind": "zip_ref",
                        "zip_bytes": zip_bytes,
                        "zip_member": info.filename,
                    })
                else:
                    # Original behavior: read file fully into memory
                    with zf.open(info) as f:
                        data = f.read()
                    processed_assets.append({
                        "name": info.filename,
                        "ext": ext,
                        "kind": "bytes",
                        "data": data,
                    })

        else:
            # Individual file
            data = uploaded_file.getvalue()
            processed_assets.append({
                "name": uploaded_file.name,
                "ext": uploaded_file.name.split(".")[-1].lower() if "." in uploaded_file.name else "",
                "kind": "bytes",
                "data": data,
            })

    return processed_assets


def asset_bytes(asset) -> Optional[bytes]:
    """Resolve bytes for an asset (either direct bytes or a zip member reference)."""
    if asset["kind"] == "bytes":
        return asset["data"]

    if asset["kind"] == "zip_ref":
        # Read that specific member on demand
        zf = zipfile.ZipFile(io.BytesIO(asset["zip_bytes"]))
        with zf.open(asset["zip_member"]) as f:
            return f.read()

    return None


# -----------------------------
# Sidebar
# -----------------------------
with st.sidebar:
    st.header("Upload")

    xlsx_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    raw_files = st.file_uploader("Upload Assets or ZIP", accept_multiple_files=True)

    safe_mode = st.toggle(
        "Safe mode (lower RAM, read ZIP files on-demand)",
        value=True,
        help="Turn ON if you have large ZIPs. Turn OFF only for small uploads (faster previews)."
    )

    if st.button("Clear Cache / Reset"):
        st.cache_resource.clear()
        st.cache_data.clear()
        st.rerun()


# -----------------------------
# Main
# -----------------------------
if xlsx_file and raw_files:
    # Load assets + Excel data
    all_assets = load_assets(raw_files, safe_mode=safe_mode)
    ad_codes, df, ad_code_col = get_ad_data_xlsx(xlsx_file)

    st.success(f"Loaded {len(all_assets)} files. Found {len(ad_codes)} Ad Codes.")

    # Fast Search
    search = st.text_input("🔍 Quick Search Ad Code", placeholder="Type to filter...")
    filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes

    for code in filtered_codes:
        # Match based on filename containing the code
        matches = [a for a in all_assets if code in a["name"]]

        if matches:
            with st.expander(f"✅ Ad {code} - {len(matches)} files found", expanded=True):
                c1, c2 = st.columns(
