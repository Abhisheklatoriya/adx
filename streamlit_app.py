import streamlit as st
import pandas as pd
import re
import zipfile
import io
import os

st.set_page_config(page_title="Ad creative matcher", layout="wide")
st.title("⚡ Ultra-Fast Ad Matcher (XLSX)")

AD_CODE_RE = re.compile(r"\b\d{8}\b")


# -----------------------------
# Safe check: openpyxl must exist
# -----------------------------
def ensure_openpyxl():
    try:
        import openpyxl  # noqa: F401
        return True
    except Exception:
        st.error(
            "Missing dependency: **openpyxl**.\n\n"
            "Fix: add this to `requirements.txt` and redeploy:\n"
            "`openpyxl==3.1.5`"
        )
        st.stop()


# -----------------------------
# Excel extraction (cached)
# -----------------------------
@st.cache_data
def get_excel_data(excel_bytes: bytes):
    ensure_openpyxl()
    df = pd.read_excel(io.BytesIO(excel_bytes), engine="openpyxl")

    # Try to detect an ad code column (best effort)
    ad_code_col = None
    for c in df.columns:
        cl = str(c).strip().lower()
        if "ad" in cl and "code" in cl:
            ad_code_col = c
            break

    # Extract unique ad codes
    codes = set()

    if ad_code_col is not None:
        series = df[ad_code_col].astype(str)
        for v in series:
            m = AD_CODE_RE.search(v)
            if m:
                codes.add(m.group(0))
    else:
        # fallback: scan entire sheet
        flat = "\n".join(df.astype(str).fillna("").values.flatten().tolist())
        codes.update(AD_CODE_RE.findall(flat))

    return df, ad_code_col, sorted(codes)


# -----------------------------
# Assets loading (cached)
# -----------------------------
@st.cache_resource
def load_assets(uploaded_files):
    processed_assets = []

    for uploaded_file in uploaded_files:
        name = uploaded_file.name

        # ZIP support (including nested folders)
        if name.lower().endswith(".zip"):
            zip_bytes = uploaded_file.getvalue()
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
                for info in z.infolist():
                    if info.is_dir() or "__MACOSX" in info.filename:
                        continue
                    with z.open(info) as f:
                        data = f.read()

                    ext = info.filename.split(".")[-1].lower() if "." in info.filename else ""
                    processed_assets.append(
                        {
                            "name": info.filename,  # keep folder path inside zip
                            "base": os.path.basename(info.filename),
                            "data": data,
                            "ext": ext,
                        }
                    )
        else:
            ext = name.split(".")[-1].lower() if "." in name else ""
            processed_assets.append(
                {
                    "name": name,
                    "base": os.path.basename(name),
                    "data": uploaded_file.getvalue(),
                    "ext": ext,
                }
            )

    return processed_assets


# -----------------------------
# Helper: preview media
# -----------------------------
def preview_asset(asset):
    ext = asset["ext"]
    data = asset["data"]

    if ext in ["mp4", "mov", "webm"]:
        st.video(data)
    elif ext in ["mp3", "wav"]:
        st.audio(data)
    elif ext in ["jpg", "jpeg", "png", "gif", "webp"]:
        st.image(data)


# -----------------------------
# Helper: get matching Excel rows for an ad code
# -----------------------------
def get_rows_for_code(df, ad_code_col, code: str):
    if df is None:
        return None

    if ad_code_col is not None and ad_code_col in df.columns:
        extracted = df[ad_code_col].astype(str).str.extract(r"(\d{8})", expand=False)
        rows = df[extracted == code]
        if len(rows) > 0:
            return rows

    # fallback: search anywhere in row
    mask = df.astype(str).apply(lambda r: r.str.contains(code, regex=False, na=False)).any(axis=1)
    rows = df[mask]
    return rows if len(rows) > 0 else None


# -----------------------------
# Sidebar upload
# -----------------------------
with st.sidebar:
    st.header("Upload")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    raw_files = st.file_uploader("Upload Assets or ZIP", accept_multiple_files=True)

    if st.button("Clear Cache / Reset"):
        st.cache_resource.clear()
        st.cache_data.clear()
        st.rerun()


if not (excel_file and raw_files):
    st.info("Please upload your **Excel** and creative assets to begin.")
    st.stop()

ensure_openpyxl()

excel_bytes = excel_file.getvalue()
df, ad_code_col, ad_codes = get_excel_data(excel_bytes)
all_assets = load_assets(raw_files)

st.success(f"Loaded {len(all_assets)} files. Found {len(ad_codes)} Ad Codes.")

# Quick search
search = st.text_input("🔍 Quick Search Ad Code", placeholder="Type to filter...")
filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes

# Render matches
for code in filtered_codes:
    # match: ad code appears in file name (folder path OR base filename)
    matches = [a for a in all_assets if code in a["name"] or code in a["base"]]

    if not matches:
        continue

    with st.expander(f"✅ Ad {code} - {len(matches)} files found", expanded=True):
        c1, c2 = st.columns([1, 1.5])

        # LEFT: Excel details
        with c1:
            st.markdown("**Ad Details (Excel):**")
            rows = get_rows_for_code(df, ad_code_col, code)
            if rows is None:
                st.warning("No matching Excel rows found for this ad code.")
            else:
                st.dataframe(rows, use_container_width=True)

        # RIGHT: Creative previews + downloads
        with c2:
            for asset in matches:
                st.caption(f"File: {asset['name']}")
                preview_asset(asset)

                st.download_button(
                    "Download",
                    data=asset["data"],
                    file_name=asset["base"],
                    key=f"dl_{asset['name']}_{code}",
                )
                st.divider()
