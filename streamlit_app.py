import streamlit as st
import pandas as pd
import re
import zipfile
import io
import os
import streamlit.components.v1 as components

# -------------------------------
# Page setup
# -------------------------------
st.set_page_config(
    page_title="Ultra-Fast Ad Matcher (XLSX)",
    layout="wide"
)

st.title("⚡ Ultra-Fast Ad Matcher (XLSX)")

# -------------------------------
# Helpers
# -------------------------------

def extract_ad_code_from_filename(name: str):
    """
    Extracts numeric ad code from filenames like:
    asset_ad_48726907_sWccqC.mp4
    """
    match = re.search(r"ad[_\-]?(\d{6,})", name.lower())
    return match.group(1) if match else None


def copy_button(text: str, label="📋 Copy details"):
    safe = (
        text.replace("\\", "\\\\")
            .replace("`", "\\`")
            .replace("$", "\\$")
            .replace("\n", "\\n")
    )

    components.html(
        f"""
        <button
            style="
                background:#f0f2f6;
                border:1px solid #d0d7de;
                padding:8px 12px;
                border-radius:8px;
                cursor:pointer;
                font-size:14px;
            "
            onclick="navigator.clipboard.writeText(`{safe}`);"
        >
            {label}
        </button>
        """,
        height=45,
    )


# -------------------------------
# Cache loaders
# -------------------------------

@st.cache_data
def load_excel(file):
    df = pd.read_excel(file)
    df.columns = [c.strip().lower() for c in df.columns]
    return df


@st.cache_resource
def load_assets(files):
    assets = []

    for f in files:
        if f.name.lower().endswith(".zip"):
            with zipfile.ZipFile(f) as z:
                for info in z.infolist():
                    if info.is_dir() or "__macosx" in info.filename.lower():
                        continue
                    with z.open(info) as zf:
                        data = zf.read()
                    assets.append({
                        "name": info.filename,
                        "data": data,
                        "ext": info.filename.split(".")[-1].lower(),
                        "ad_code": extract_ad_code_from_filename(info.filename)
                    })
        else:
            assets.append({
                "name": f.name,
                "data": f.getvalue(),
                "ext": f.name.split(".")[-1].lower(),
                "ad_code": extract_ad_code_from_filename(f.name)
            })

    return assets


# -------------------------------
# Sidebar
# -------------------------------

with st.sidebar:
    st.header("Upload")

    excel_file = st.file_uploader(
        "Upload Excel (.xlsx)",
        type=["xlsx"]
    )

    asset_files = st.file_uploader(
        "Upload Assets or ZIP",
        accept_multiple_files=True
    )

    if st.button("Clear Cache / Reset"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.rerun()


# -------------------------------
# Main logic
# -------------------------------

if excel_file and asset_files:

    df = load_excel(excel_file)
    assets = load_assets(asset_files)

    # Find ad code column
    ad_code_col = next(
        (c for c in df.columns if "ad" in c and "code" in c),
        None
    )

    if not ad_code_col:
        st.error("No Ad Code column found in Excel.")
        st.stop()

    df[ad_code_col] = df[ad_code_col].astype(str)

    ad_codes = sorted(df[ad_code_col].unique())

    st.success(f"Loaded {len(asset_files)} files. Found {len(ad_codes)} Ad Codes.")

    search = st.text_input("🔍 Quick Search Ad Code", placeholder="Type to filter...")
    visible_codes = [c for c in ad_codes if search in c] if search else ad_codes

    for ad_code in visible_codes:
        matched_assets = [
            a for a in assets
            if a["ad_code"] == ad_code
        ]

        if not matched_assets:
            continue

        with st.expander(f"✅ Ad Code {ad_code} — {len(matched_assets)} asset(s)", expanded=True):

            ad_rows = df[df[ad_code_col] == ad_code]

            left, right = st.columns([1, 1.5])

            # -------------------------------
            # LEFT: Ad details
            # -------------------------------
            with left:
                st.markdown("### Ad Details")

                st.dataframe(ad_rows, use_container_width=True)

                details_text = "\n".join(
                    f"{col}: {ad_rows.iloc[0][col]}"
                    for col in ad_rows.columns
                )

                st.text_area(
                    "Copy-friendly details",
                    details_text,
                    height=260
                )

                copy_button(details_text)

            # -------------------------------
            # RIGHT: Creative preview
            # -------------------------------
            with right:
                for asset in matched_assets:
                    st.markdown(f"**File:** {asset['name']}")

                    if asset["ext"] in ["mp4", "mov", "webm"]:
                        st.video(asset["data"])
                    elif asset["ext"] in ["jpg", "jpeg", "png", "gif"]:
                        st.image(asset["data"])
                    elif asset["ext"] in ["mp3", "wav"]:
                        st.audio(asset["data"])

                    st.download_button(
                        "⬇ Download asset",
                        data=asset["data"],
                        file_name=os.path.basename(asset["name"]),
                        key=f"dl_{asset['name']}"
                    )

                    st.divider()

    # -------------------------------
    # SAVE SESSION AS ZIP
    # -------------------------------
    st.markdown("---")
    st.subheader("💾 Save Session")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("ad_details.xlsx", excel_file.getvalue())
        for a in assets:
            z.writestr(a["name"], a["data"])

    st.download_button(
        "⬇ Download entire session (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="ad_matcher_session.zip",
        mime="application/zip"
    )

else:
    st.info("Upload an Excel file and creatives to begin.")
