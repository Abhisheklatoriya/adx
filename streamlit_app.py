import streamlit as st
import pandas as pd
import re
import zipfile
import io
import os
from datetime import datetime

# -----------------------------
# Config
# -----------------------------
st.set_page_config(page_title="Ad creative matcher", layout="wide")
st.title("⚡ Ultra-Fast Ad Matcher (XLSX) + Session Save/Restore")

AD_CODE_RE = re.compile(r"\b\d{8}\b")


# -----------------------------
# Excel extraction (cached)
# -----------------------------
@st.cache_data
def get_ad_data(excel_file):
    """
    Returns:
      ad_codes: sorted unique list of 8-digit ad codes found (prefer Ad Code col if detected)
      df: original dataframe
      ad_code_col: detected column name containing ad code (best effort)
      excel_bytes: raw bytes of the uploaded excel (for session export)
    """
    excel_bytes = excel_file.getvalue()

    # IMPORTANT: requires openpyxl installed on Streamlit Cloud
    df = pd.read_excel(io.BytesIO(excel_bytes), engine="openpyxl")

    # detect ad code column
    ad_code_col = None
    for c in df.columns:
        cl = str(c).strip().lower()
        if "ad" in cl and "code" in cl:
            ad_code_col = c
            break

    codes = set()
    if ad_code_col is not None:
        s = df[ad_code_col].astype(str)
        for v in s:
            m = AD_CODE_RE.search(v)
            if m:
                codes.add(m.group(0))
    else:
        # scan all cells
        flat = "\n".join(df.astype(str).fillna("").values.flatten().tolist())
        codes.update(AD_CODE_RE.findall(flat))

    ad_codes = sorted(codes)
    return ad_codes, df, ad_code_col, excel_bytes


# -----------------------------
# Assets loading (cached)
# -----------------------------
@st.cache_resource
def load_assets(uploaded_files):
    """
    Loads all assets into memory (like your original code).
    Supports:
      - individual files
      - zip (extracts all files except __MACOSX + directories)
    """
    processed_assets = []
    for uploaded_file in uploaded_files:
        if uploaded_file.name.lower().endswith(".zip"):
            zip_bytes = uploaded_file.getvalue()
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
                for info in z.infolist():
                    if info.is_dir() or "__MACOSX" in info.filename:
                        continue
                    with z.open(info) as f:
                        data = f.read()
                    ext = info.filename.split(".")[-1].lower() if "." in info.filename else ""
                    processed_assets.append({"name": info.filename, "data": data, "ext": ext})
        else:
            ext = uploaded_file.name.split(".")[-1].lower() if "." in uploaded_file.name else ""
            processed_assets.append({"name": uploaded_file.name, "data": uploaded_file.getvalue(), "ext": ext})
    return processed_assets


# -----------------------------
# Restore Session ZIP
# -----------------------------
def restore_session_zip(session_zip_file):
    """
    Reads a previously exported session zip and returns:
      - excel_bytes (optional)
      - df, ad_codes, ad_code_col (derived from excel_bytes if present)
      - assets list with bytes (from creatives/...)
      - manifest dataframe (if present)
    """
    zbytes = session_zip_file.getvalue()
    assets = []
    manifest_df = None
    excel_bytes = None

    with zipfile.ZipFile(io.BytesIO(zbytes)) as z:
        # load excel if present
        for name in z.namelist():
            lower = name.lower()
            if lower.endswith(".xlsx") and ("session.xlsx" in lower or "/session.xlsx" in lower or "excel" in lower):
                excel_bytes = z.read(name)
                break

        # load manifest if present
        for name in z.namelist():
            if name.lower().endswith("manifest.csv"):
                manifest_df = pd.read_csv(io.BytesIO(z.read(name)))
                break

        # load creatives
        for name in z.namelist():
            if name.endswith("/") or "__MACOSX" in name:
                continue
            # Only load files under creatives/
            if name.startswith("creatives/"):
                data = z.read(name)
                ext = name.split(".")[-1].lower() if "." in name else ""
                assets.append({"name": name.replace("creatives/", "", 1), "data": data, "ext": ext})

    # derive df/ad_codes if excel included
    df = None
    ad_codes = []
    ad_code_col = None
    if excel_bytes:
        # reuse the same logic (without caching)
        df = pd.read_excel(io.BytesIO(excel_bytes), engine="openpyxl")

        for c in df.columns:
            cl = str(c).strip().lower()
            if "ad" in cl and "code" in cl:
                ad_code_col = c
                break

        codes = set()
        if ad_code_col is not None:
            s = df[ad_code_col].astype(str)
            for v in s:
                m = AD_CODE_RE.search(v)
                if m:
                    codes.add(m.group(0))
        else:
            flat = "\n".join(df.astype(str).fillna("").values.flatten().tolist())
            codes.update(AD_CODE_RE.findall(flat))

        ad_codes = sorted(codes)

    return excel_bytes, df, ad_codes, ad_code_col, assets, manifest_df


# -----------------------------
# Helper: show media preview
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
# Helper: extract matching rows for left pane
# -----------------------------
def rows_for_code(df, ad_code_col, code):
    if df is None:
        return None

    if ad_code_col is not None and ad_code_col in df.columns:
        col_as_code = df[ad_code_col].astype(str).str.extract(r"(\d{8})", expand=False)
        rows = df[col_as_code == code]
        if len(rows) > 0:
            return rows

    # fallback: match code anywhere
    mask = df.astype(str).apply(lambda r: r.str.contains(code, regex=False, na=False)).any(axis=1)
    rows = df[mask]
    return rows if len(rows) > 0 else df.head(5)


# -----------------------------
# Export Session ZIP
# -----------------------------
def build_session_zip(excel_bytes, df, ad_code_col, matches_map, assets):
    """
    Creates a zip with:
      - session.xlsx
      - manifest.csv
      - creatives/<ad_code>/<filename>
    matches_map: dict ad_code -> list of asset dicts
    assets: all assets list (not strictly needed, but here if you want later)
    """
    out = io.BytesIO()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Build manifest rows
    manifest_rows = []
    for code, matched_assets in matches_map.items():
        rows = rows_for_code(df, ad_code_col, code)
        # store excel row indices for reference (best effort)
        row_ids = []
        if rows is not None and len(rows) > 0:
            row_ids = rows.index.tolist()

        for a in matched_assets:
            manifest_rows.append({
                "ad_code": code,
                "creative_name": a["name"],
                "excel_row_indices": ",".join(map(str, row_ids)) if row_ids else ""
            })

    manifest_df = pd.DataFrame(manifest_rows)

    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as z:
        # Excel
        if excel_bytes:
            z.writestr("session.xlsx", excel_bytes)

        # Manifest
        z.writestr("manifest.csv", manifest_df.to_csv(index=False))

        # Creatives organized
        for code, matched_assets in matches_map.items():
            for a in matched_assets:
                # Keep only the base filename inside the ad folder
                base = os.path.basename(a["name"])
                zip_path = f"creatives/{code}/{base}"
                z.writestr(zip_path, a["data"])

    out.seek(0)
    filename = f"ad_match_session_{timestamp}.zip"
    return out.getvalue(), filename


# -----------------------------
# Sidebar UI
# -----------------------------
with st.sidebar:
    st.header("Session")
    session_zip = st.file_uploader("Restore Session ZIP (optional)", type=["zip"])

    st.divider()
    st.header("Upload (New Session)")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    raw_files = st.file_uploader("Upload Assets or ZIP", accept_multiple_files=True)

    if st.button("Clear Cache / Reset"):
        st.cache_resource.clear()
        st.cache_data.clear()
        st.rerun()


# -----------------------------
# Main logic: Restore OR New upload
# -----------------------------
restored = False
excel_bytes = None
df = None
ad_codes = []
ad_code_col = None
all_assets = []
manifest_df = None

if session_zip:
    # RESTORE MODE
    restored = True
    excel_bytes, df, ad_codes, ad_code_col, all_assets, manifest_df = restore_session_zip(session_zip)

    st.success(f"✅ Restored session. Loaded {len(all_assets)} creatives. Found {len(ad_codes)} Ad Codes.")
    if manifest_df is not None:
        with st.expander("View manifest.csv", expanded=False):
            st.dataframe(manifest_df, use_container_width=True)

elif excel_file and raw_files:
    # NEW SESSION MODE
    ad_codes, df, ad_code_col, excel_bytes = get_ad_data(excel_file)
    all_assets = load_assets(raw_files)
    st.success(f"Loaded {len(all_assets)} files. Found {len(ad_codes)} Ad Codes.")
else:
    st.info("Restore a session ZIP, or upload your Excel + creatives to begin.")
    st.stop()


# -----------------------------
# Search + Render (same UX as original)
# -----------------------------
search = st.text_input("🔍 Quick Search Ad Code", placeholder="Type to filter...")
filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes

# Build matches map once (used for export + UI)
matches_map = {}
for code in filtered_codes:
    matches = [a for a in all_assets if code in a["name"]]
    if matches:
        matches_map[code] = matches

# Export session ZIP (only if we have something matched)
if len(matches_map) > 0:
    st.divider()
    cexp1, cexp2 = st.columns([1, 2])
    with cexp1:
        st.markdown("### 💾 Save / Share this session")
        if not excel_bytes:
            st.warning("Excel bytes not found (restore ZIP may be missing session.xlsx). Export will still include creatives + manifest.")
    with cexp2:
        session_zip_bytes, session_zip_name = build_session_zip(
            excel_bytes=excel_bytes,
            df=df,
            ad_code_col=ad_code_col,
            matches_map=matches_map,
            assets=all_assets,
        )
        st.download_button(
            "⬇️ Export Session ZIP",
            data=session_zip_bytes,
            file_name=session_zip_name,
            mime="application/zip",
            key="export_session_zip",
            help="Share this ZIP with anyone. They can restore it without re-uploading everything."
        )

st.divider()

# Render each matched code like your original app
for code in filtered_codes:
    if code not in matches_map:
        continue

    matches = matches_map[code]
    with st.expander(f"✅ Ad {code} - {len(matches)} files found", expanded=True):
        c1, c2 = st.columns([1, 1.5])

        # LEFT: Ad Specs
        with c1:
            st.markdown("**Ad Specs:**")
            rows = rows_for_code(df, ad_code_col, code)
            if rows is None:
                st.info("No Excel data available in this session.")
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
                    file_name=os.path.basename(asset["name"]),
                    key=f"dl_{asset['name']}_{code}",
                )
                st.divider()
