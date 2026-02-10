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
st.set_page_config(page_title="Ultra-Fast Ad Matcher (XLSX)", layout="wide")
st.title("⚡ Ultra-Fast Ad Matcher (XLSX)")

# -------------------------------
# Guard: openpyxl must exist for read_excel
# -------------------------------
try:
    import openpyxl  # noqa: F401
except Exception:
    st.error(
        "Your deployment is missing **openpyxl**, so Excel files can't be read yet.\n\n"
        "Fix:\n"
        "1) In GitHub, set `requirements.txt` to:\n"
        "   - streamlit==1.54.0\n"
        "   - pandas==2.3.3\n"
        "   - openpyxl==3.1.5\n"
        "2) Commit + push\n"
        "3) In Streamlit Cloud → Manage app → **Reboot**\n\n"
        "Once openpyxl installs, refresh and it will work."
    )
    st.stop()

# -------------------------------
# Helpers
# -------------------------------
def extract_ad_code_from_filename(name: str):
    """
    Extract numeric ad code from filenames like:
    asset_ad_48734339_WAUOgJ.mp4
    asset-ad-48734339.mp4
    """
    n = name.lower()
    m = re.search(r"(?:^|[^a-z0-9])ad[_\-]?(\d{6,})(?:[^0-9]|$)", n)
    return m.group(1) if m else None


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


@st.cache_data
def load_excel(file):
    df = pd.read_excel(file, engine="openpyxl")
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


@st.cache_resource
def load_assets(files):
    assets = []

    for f in files:
        if f.name.lower().endswith(".zip"):
            with zipfile.ZipFile(f) as z:
                for info in z.infolist():
                    if info.is_dir():
                        continue
                    if "__macosx" in info.filename.lower():
                        continue
                    with z.open(info) as zf:
                        data = zf.read()

                    assets.append({
                        "name": info.filename,
                        "data": data,
                        "ext": info.filename.split(".")[-1].lower() if "." in info.filename else "",
                        "ad_code": extract_ad_code_from_filename(info.filename),
                    })
        else:
            assets.append({
                "name": f.name,
                "data": f.getvalue(),
                "ext": f.name.split(".")[-1].lower() if "." in f.name else "",
                "ad_code": extract_ad_code_from_filename(f.name),
            })

    return assets


# -------------------------------
# Sidebar
# -------------------------------
with st.sidebar:
    st.header("Upload")

    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    asset_files = st.file_uploader("Upload Assets or ZIP", accept_multiple_files=True)

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

    # Find an "Ad Code" column robustly
    ad_code_col = None
    for c in df.columns:
        c_norm = c.replace("_", " ").strip()
        if "ad code" in c_norm or (("ad" in c_norm) and ("code" in c_norm)):
            ad_code_col = c
            break

    if not ad_code_col:
        st.error("Excel load error: Could not find an Ad Code column (expected something like 'Ad Code').")
        st.stop()

    df[ad_code_col] = df[ad_code_col].astype(str).str.strip()

    ad_codes = sorted([c for c in df[ad_code_col].unique() if c and c.lower() != "nan"])

    st.success(f"Loaded {len(asset_files)} file(s). Found {len(ad_codes)} Ad Code(s).")

    search = st.text_input("🔍 Quick Search Ad Code", placeholder="Type to filter...")
    visible_codes = [c for c in ad_codes if search in c] if search else ad_codes

    # Build quick lookup from assets by extracted ad_code
    assets_by_code = {}
    for a in assets:
        if not a["ad_code"]:
            continue
        assets_by_code.setdefault(a["ad_code"], []).append(a)

    shown_any = False

    for ad_code in visible_codes:
        matched_assets = assets_by_code.get(ad_code, [])
        if not matched_assets:
            continue

        shown_any = True

        with st.expander(f"✅ Ad Code {ad_code} — {len(matched_assets)} asset(s)", expanded=True):

            ad_rows = df[df[ad_code_col] == ad_code]

            left, right = st.columns([1, 1.5])

            # LEFT: ad details + copy
            with left:
                st.markdown("### Ad Details")
                st.dataframe(ad_rows, use_container_width=True)

                row0 = ad_rows.iloc[0].to_dict() if len(ad_rows) else {"ad_code": ad_code}
                details_text = "\n".join([f"{k}: {row0.get(k)}" for k in ad_rows.columns])

                st.text_area("Copy-friendly details", details_text, height=260)
                copy_button(details_text, label="📋 Copy details")

            # RIGHT: previews + download buttons
            with right:
                for asset in matched_assets:
                    st.markdown(f"**File:** {asset['name']}")

                    ext = asset["ext"]
                    if ext in ["mp4", "mov", "webm"]:
                        st.video(asset["data"])
                    elif ext in ["jpg", "jpeg", "png", "gif"]:
                        st.image(asset["data"])
                    elif ext in ["mp3", "wav"]:
                        st.audio(asset["data"])
                    else:
                        st.info(f"No inline preview for .{ext} — download to view.")

                    st.download_button(
                        "⬇ Download asset",
                        data=asset["data"],
                        file_name=os.path.basename(asset["name"]),
                        key=f"dl_{asset['name']}",
                    )
                    st.divider()

    if not shown_any:
        st.warning(
            "No matches found.\n\n"
            "This usually means either:\n"
            "- The Excel ad codes are not the same digits as the filenames, OR\n"
            "- Filenames don’t contain `ad_<digits>` pattern.\n\n"
            "Example supported filename:\n"
            "`asset_ad_48734339_WAUOgJ.mp4`"
        )

    # SAVE SESSION ZIP
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
        mime="application/zip",
    )

else:
    st.info("Upload an Excel file and creatives to begin.")
