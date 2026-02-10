import streamlit as st
import pandas as pd
import re
import zipfile
import io

# -------------------------------
# App Config
# -------------------------------
st.set_page_config(page_title="Ad Creative Matcher", layout="wide")
st.title("⚡ Ultra-Fast Ad Matcher (Excel Version)")

# -------------------------------
# Load Excel (cached)
# -------------------------------
@st.cache_data
def load_excel(file):
    df = pd.read_excel(file)

    # Normalize column names
    df.columns = [c.strip().lower() for c in df.columns]

    # Try to auto-detect ad code column
    ad_code_col = next(
        (c for c in df.columns if "ad" in c and "code" in c),
        None
    )

    if not ad_code_col:
        raise ValueError("No Ad Code column found")

    # Extract only valid 8-digit ad codes
    df[ad_code_col] = df[ad_code_col].astype(str)
    df = df[df[ad_code_col].str.match(r"^\d{8}$")]

    return df, ad_code_col


# -------------------------------
# Load Assets (cached in memory)
# -------------------------------
@st.cache_resource
def load_assets(uploaded_files):
    assets = []

    for uploaded_file in uploaded_files:
        if uploaded_file.name.lower().endswith(".zip"):
            with zipfile.ZipFile(uploaded_file) as z:
                for info in z.infolist():
                    if info.is_dir() or "__MACOSX" in info.filename:
                        continue
                    with z.open(info) as f:
                        assets.append({
                            "name": info.filename,
                            "data": f.read(),
                            "ext": info.filename.split(".")[-1].lower()
                        })
        else:
            assets.append({
                "name": uploaded_file.name,
                "data": uploaded_file.getvalue(),
                "ext": uploaded_file.name.split(".")[-1].lower()
            })

    return assets


# -------------------------------
# Sidebar
# -------------------------------
with st.sidebar:
    st.header("Upload")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    raw_files = st.file_uploader(
        "Upload Assets or ZIPs",
        accept_multiple_files=True
    )

    if st.button("Clear Cache / Reset"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.rerun()

# -------------------------------
# Main Logic
# -------------------------------
if excel_file and raw_files:
    df, ad_code_col = load_excel(excel_file)
    assets = load_assets(raw_files)

    st.success(
        f"Loaded {len(df)} ads · {len(assets)} creative files"
    )

    # Quick search
    search = st.text_input(
        "🔍 Search Ad Code",
        placeholder="Type partial or full code"
    )

    ad_codes = df[ad_code_col].tolist()
    filtered_codes = [
        c for c in ad_codes if search in c
    ] if search else ad_codes

    for ad_code in filtered_codes:
        matches = [
            a for a in assets if ad_code in a["name"]
        ]

        if not matches:
            continue

        with st.expander(
            f"✅ Ad {ad_code} — {len(matches)} files",
            expanded=True
        ):
            c1, c2 = st.columns([1, 1.6])

            # -----------------------
            # Specs column
            # -----------------------
            with c1:
                st.markdown("**Ad Specs**")
                row = df[df[ad_code_col] == ad_code]
                st.dataframe(row, use_container_width=True)

            # -----------------------
            # Creatives column
            # -----------------------
            with c2:
                for asset in matches:
                    st.caption(asset["name"])

                    if asset["ext"] in ["mp4", "mov", "webm"]:
                        st.video(asset["data"])
                    elif asset["ext"] in ["mp3", "wav"]:
                        st.audio(asset["data"])
                    elif asset["ext"] in ["jpg", "jpeg", "png", "gif"]:
                        st.image(asset["data"])

                    st.download_button(
                        "Download",
                        data=asset["data"],
                        file_name=asset["name"],
                        key=f"{ad_code}_{asset['name']}"
                    )
                    st.divider()

else:
    st.info("Upload your Excel ad list and creative assets to begin.")
