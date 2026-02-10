# streamlit_app.py
import io
import os
import re
import time
import json
import hashlib
import zipfile
import tempfile
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# -------------------------------
# Config
# -------------------------------
st.set_page_config(page_title="Ultra-Fast Ad Matcher (XLSX)", layout="wide")

APP_TITLE = "⚡ Ultra-Fast Ad Matcher (XLSX)"
SESSION_VERSION = "1.0"

# Asset filename pattern examples:
# - asset_ad_48734339_WAUOgJ.mp4
# We'll extract the numeric ad code after "asset_ad_" by default.
ASSET_ADCODE_REGEX = re.compile(r"(?:^|[_\-\s])asset[_\-\s]?ad[_\-\s]?(\d+)", re.IGNORECASE)

# If you also see variants like "ad_48734339" without "asset", enable fallback:
FALLBACK_ADCODE_REGEX = re.compile(r"(?:^|[_\-\s])ad[_\-\s]?(\d+)", re.IGNORECASE)

ALLOWED_ASSET_EXTS = {
    ".mp4", ".mov", ".m4v", ".webm",
    ".jpg", ".jpeg", ".png", ".gif", ".webp",
    ".pdf",
    ".zip",  # handled specially
    ".html", ".htm",
    ".mp3", ".wav"
}


# -------------------------------
# Helpers
# -------------------------------
def _safe_filename(name: str) -> str:
    name = os.path.basename(name)
    name = re.sub(r"[^a-zA-Z0-9._\- ]+", "_", name).strip()
    return name[:200] if name else "file"


def extract_ad_code_from_asset_filename(filename: str) -> Optional[str]:
    """
    Extract numeric ad code from filename using expected pattern.
    Returns the digits as string (e.g., "48734339") or None.
    """
    base = os.path.basename(filename)
    m = ASSET_ADCODE_REGEX.search(base)
    if m:
        return m.group(1)

    m2 = FALLBACK_ADCODE_REGEX.search(base)
    if m2:
        return m2.group(1)

    return None


def normalize_str(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())


@st.cache_data
def load_excel(file) -> Tuple[pd.DataFrame, Optional[str]]:
    """
    Loads excel and tries to auto-detect ad-code-like column.
    If not found, returns (df, None) so UI can ask user to pick.
    """
    df = pd.read_excel(file, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    norm_map = {c: normalize_str(c) for c in df.columns}

    target_norms = {
        "adcode", "adcodes", "ad_code", "ad code",
        "adid", "adids", "ad id",
        "firstrunadid", "first_run_ad_id", "first run ad id",
        "creativeid", "creative_id", "creative id",
    }
    target_norms = {normalize_str(t) for t in target_norms}

    best_col = None
    for col, n in norm_map.items():
        if n in target_norms:
            best_col = col
            break

    return df, best_col


@dataclass
class AssetItem:
    name: str
    ext: str
    ad_code: Optional[str]
    bytes_data: bytes


def _read_uploaded_file_bytes(uploaded) -> bytes:
    return uploaded.getvalue()


def _zip_list_files(zip_bytes: bytes) -> List[str]:
    names = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            names.append(info.filename)
    return names


def load_assets(uploaded_files: List) -> List[AssetItem]:
    """
    Accepts files and/or zip(s). If zip, extracts all files inside.
    Returns list of AssetItem with bytes stored (good for preview + downloads).
    """
    assets: List[AssetItem] = []

    for uf in uploaded_files:
        fname = _safe_filename(uf.name)
        ext = os.path.splitext(fname)[1].lower()

        raw = _read_uploaded_file_bytes(uf)

        if ext == ".zip":
            # Extract zip contents
            with zipfile.ZipFile(io.BytesIO(raw), "r") as zf:
                for info in zf.infolist():
                    if info.is_dir():
                        continue
                    inner_name = _safe_filename(info.filename)
                    inner_ext = os.path.splitext(inner_name)[1].lower()
                    if inner_ext and inner_ext not in ALLOWED_ASSET_EXTS:
                        # still allow unknown ext; we'll just not preview
                        pass
                    data = zf.read(info.filename)
                    ad_code = extract_ad_code_from_asset_filename(inner_name)
                    assets.append(AssetItem(name=inner_name, ext=inner_ext, ad_code=ad_code, bytes_data=data))
        else:
            ad_code = extract_ad_code_from_asset_filename(fname)
            assets.append(AssetItem(name=fname, ext=ext, ad_code=ad_code, bytes_data=raw))

    return assets


def build_matches(df: pd.DataFrame, ad_code_col: str, assets: List[AssetItem]) -> Tuple[pd.DataFrame, Dict[str, List[AssetItem]]]:
    """
    Returns:
      - df_clean: df with a normalized ad_code_key column (string)
      - match_map: {ad_code_key: [AssetItem,...]}
    Matching logic:
      - Extract numeric ad code from asset filename, compare to numeric portion from excel ad_code_col
      - Handles cases like "48726907D" in Excel by extracting leading digits
    """
    df = df.copy()

    # Create key from excel: extract first digit run in the cell
    def excel_to_key(x) -> Optional[str]:
        if pd.isna(x):
            return None
        s = str(x).strip()
        m = re.search(r"(\d+)", s)
        return m.group(1) if m else None

    df["__ad_code_key__"] = df[ad_code_col].apply(excel_to_key)

    # Build map from assets
    match_map: Dict[str, List[AssetItem]] = {}
    for a in assets:
        if not a.ad_code:
            continue
        match_map.setdefault(a.ad_code, []).append(a)

    return df, match_map


def preview_asset(asset: AssetItem):
    """
    Show a preview for common media.
    """
    st.caption(f"File: {asset.name}")
    ext = asset.ext.lower()

    if ext in [".mp4", ".mov", ".m4v", ".webm"]:
        st.video(asset.bytes_data)
    elif ext in [".mp3", ".wav"]:
        st.audio(asset.bytes_data)
    elif ext in [".jpg", ".jpeg", ".png", ".gif", ".webp"]:
        st.image(asset.bytes_data, use_container_width=True)
    elif ext == ".pdf":
        # Streamlit doesn't have native PDF viewer; provide download.
        st.info("PDF preview isn't native in Streamlit; use download below.")
    elif ext in [".html", ".htm"]:
        try:
            html = asset.bytes_data.decode("utf-8", errors="ignore")
            st.components.v1.html(html, height=500, scrolling=True)
        except Exception:
            st.info("Couldn't render HTML preview; use download.")
    else:
        st.info("No preview for this file type. Use download.")


def copy_block(text: str, key: str):
    """
    Renders a textarea + a Copy button using a small JS snippet.
    """
    st.text_area("Copy-friendly details", value=text, height=260, key=key)

    # Copy button via JS
    # NOTE: Streamlit doesn't have a native clipboard API.
    # This works in most browsers.
    btn_key = f"{key}_copybtn"
    if st.button("Copy", key=btn_key):
        escaped = text.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
        st.components.v1.html(
            f"""
            <script>
            const text = `{escaped}`;
            navigator.clipboard.writeText(text).then(() => {{
              const el = window.parent.document.querySelector('div[data-testid="stToast"]');
            }});
            </script>
            """,
            height=0,
        )
        st.success("Copied to clipboard.")


def make_session_zip_bytes(
    excel_name: str,
    excel_bytes: bytes,
    assets: List[AssetItem],
    meta: dict
) -> bytes:
    """
    Creates a 'session bundle' zip containing:
      - session/meta.json
      - inputs/excel/<excel file>
      - inputs/assets/<each asset>
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("session/meta.json", json.dumps(meta, indent=2))

        zf.writestr(f"inputs/excel/{_safe_filename(excel_name)}", excel_bytes)

        for a in assets:
            zf.writestr(f"inputs/assets/{_safe_filename(a.name)}", a.bytes_data)

    return buf.getvalue()


def load_session_zip(zip_bytes: bytes) -> Tuple[Optional[bytes], Optional[str], List[AssetItem], dict]:
    """
    Loads a previously saved session zip.
    Returns: (excel_bytes, excel_name, assets, meta)
    """
    excel_bytes = None
    excel_name = None
    assets: List[AssetItem] = []
    meta = {}

    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
        # meta
        if "session/meta.json" in zf.namelist():
            meta = json.loads(zf.read("session/meta.json").decode("utf-8", errors="ignore"))

        # excel: take first file under inputs/excel/
        excel_candidates = [n for n in zf.namelist() if n.startswith("inputs/excel/") and not n.endswith("/")]
        if excel_candidates:
            excel_name = os.path.basename(excel_candidates[0])
            excel_bytes = zf.read(excel_candidates[0])

        # assets
        asset_candidates = [n for n in zf.namelist() if n.startswith("inputs/assets/") and not n.endswith("/")]
        for n in asset_candidates:
            name = os.path.basename(n)
            ext = os.path.splitext(name)[1].lower()
            data = zf.read(n)
            ad_code = extract_ad_code_from_asset_filename(name)
            assets.append(AssetItem(name=name, ext=ext, ad_code=ad_code, bytes_data=data))

    return excel_bytes, excel_name, assets, meta


def make_assets_zip_for_ad(ad_code_key: str, assets: List[AssetItem]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for a in assets:
            zf.writestr(_safe_filename(a.name), a.bytes_data)
    return buf.getvalue()


# -------------------------------
# Sidebar: Uploads + Session Save/Load
# -------------------------------
st.title(APP_TITLE)

with st.sidebar:
    st.header("Upload")

    # Session Load (optional)
    st.subheader("Load saved session (ZIP)")
    session_zip = st.file_uploader("Upload session ZIP", type=["zip"], key="session_zip")
    if session_zip is not None:
        try:
            zip_bytes = session_zip.getvalue()
            exb, exn, saved_assets, meta = load_session_zip(zip_bytes)
            if exb is None:
                st.error("Session ZIP didn't contain an Excel file.")
            else:
                st.session_state["saved_excel_bytes"] = exb
                st.session_state["saved_excel_name"] = exn or "session.xlsx"
                st.session_state["saved_assets"] = saved_assets
                st.session_state["loaded_from_session_zip"] = True
                st.success("Session loaded. Scroll to main area.")
        except Exception as e:
            st.error(f"Failed to load session ZIP: {e}")

    st.divider()

    # Excel upload OR use session loaded excel
    st.subheader("Upload Excel (.xlsx)")
    excel_file = st.file_uploader("Excel file", type=["xlsx"], key="excel_uploader")

    # Assets upload OR use session loaded assets
    st.subheader("Upload Assets or ZIP")
    raw_files = st.file_uploader(
        "Assets (multiple) or ZIP",
        accept_multiple_files=True,
        type=list({ext.replace(".", "") for ext in ALLOWED_ASSET_EXTS if ext != ".zip"}) + ["zip"],
        key="assets_uploader",
    )

    st.divider()
    if st.button("Clear Cache / Reset"):
        st.cache_data.clear()
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.success("Cleared. Refresh the page if needed.")


# -------------------------------
# Acquire inputs (either direct upload or loaded session)
# -------------------------------
excel_bytes = None
excel_name = None
assets: List[AssetItem] = []

# Prefer direct uploads; otherwise session zip-loaded values
if excel_file is not None:
    excel_bytes = excel_file.getvalue()
    excel_name = excel_file.name
elif st.session_state.get("saved_excel_bytes"):
    excel_bytes = st.session_state["saved_excel_bytes"]
    excel_name = st.session_state.get("saved_excel_name", "session.xlsx")

if raw_files:
    assets = load_assets(raw_files)
elif st.session_state.get("saved_assets"):
    assets = st.session_state["saved_assets"]


# -------------------------------
# Main logic
# -------------------------------
if excel_bytes and assets:
    excel_buf = io.BytesIO(excel_bytes)

    df, detected_col = load_excel(excel_buf)

    # UI: choose ad code column if not detected
    ad_code_col = detected_col
    if ad_code_col is None:
        st.warning("Excel load: couldn't auto-detect the Ad Code column. Please select it.")
        ad_code_col = st.selectbox("Select the Ad Code column", options=list(df.columns), key="ad_code_picker")

    # Build matches
    df_clean, match_map = build_matches(df, ad_code_col, assets)

    # Summary
    available_adcodes = sorted([k for k in df_clean["__ad_code_key__"].dropna().unique().tolist() if k])
    st.success(f"Loaded {len(assets)} assets. Found {len(available_adcodes)} Ad Codes in Excel.")

    # Save Session ZIP (download)
    st.subheader("Session")
    meta = {
        "version": SESSION_VERSION,
        "created_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "excel_name": excel_name,
        "asset_count": len(assets),
        "notes": "Upload this ZIP later to restore the same session without re-uploading files one-by-one.",
    }
    session_zip_bytes = make_session_zip_bytes(excel_name, excel_bytes, assets, meta)
    st.download_button(
        "Download session as ZIP",
        data=session_zip_bytes,
        file_name=f"ad_match_session_{int(time.time())}.zip",
        mime="application/zip",
        use_container_width=True,
    )

    st.divider()

    # Quick search filter
    st.markdown("🔎 **Quick Search Ad Code**")
    q = st.text_input("Type to filter…", value="", key="search_box").strip()

    # Build a small table for browsing
    browse = df_clean.copy()
    browse["matched_assets"] = browse["__ad_code_key__"].apply(lambda k: len(match_map.get(k, [])) if k else 0)

    # Filter
    if q:
        browse = browse[browse["__ad_code_key__"].astype(str).str.contains(q, na=False)]

    # Show only a few columns up front, but keep original df available
    show_cols = []
    # put ad code col + a few common columns if present
    preferred = [ad_code_col, "media", "first run date", "last run date", "brand", "headline", "description"]
    existing = {c.lower(): c for c in browse.columns}
    for p in preferred:
        if p is None:
            continue
        key = p.lower()
        if p in browse.columns:
            show_cols.append(p)
        elif key in existing:
            show_cols.append(existing[key])

    # Always include normalized key + matched count
    if "__ad_code_key__" not in show_cols:
        show_cols.insert(0, "__ad_code_key__")
    if "matched_assets" not in show_cols:
        show_cols.insert(1, "matched_assets")

    st.dataframe(
        browse[show_cols].rename(columns={"__ad_code_key__": "ad code key"}),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()

    # Selection
    left, right = st.columns([0.55, 0.45], gap="large")

    with left:
        st.subheader("Pick an Ad Code to preview")
        selected = st.selectbox(
            "Ad Code (numeric key)",
            options=available_adcodes,
            index=0 if available_adcodes else None,
            key="adcode_select",
        )

        if selected:
            # Show all rows matching selected
            rows = df_clean[df_clean["__ad_code_key__"] == selected].copy()

            # Make a copy-friendly detail block (first row)
            r0 = rows.iloc[0].to_dict()

            # Format details nicely
            lines = []
            # Put the original ad code cell too
            lines.append(f"ad code: {r0.get(ad_code_col)}")
            lines.append(f"ad code key: {selected}")

            # Add a few useful fields if present
            for col in df_clean.columns:
                if col.startswith("__"):
                    continue
                if col == ad_code_col:
                    continue
                val = r0.get(col)
                if pd.isna(val):
                    continue
                # reduce noise: skip huge strings
                s = str(val)
                if len(s) > 250:
                    s = s[:250] + "…"
                lines.append(f"{col}: {s}")

            detail_text = "\n".join(lines)

            # Table of matching rows
            st.markdown("**Matched Excel rows**")
            st.dataframe(rows.drop(columns=["__ad_code_key__"]), use_container_width=True, hide_index=True)

            # Copy-friendly details + Copy button
            copy_block(detail_text, key=f"copy_{selected}")

    with right:
        if selected:
            matched_assets = match_map.get(selected, [])

            st.subheader("Matched asset preview")
            if not matched_assets:
                st.warning("No assets matched this Ad Code key.")
            else:
                # if multiple, let user choose which to preview
                if len(matched_assets) > 1:
                    asset_names = [a.name for a in matched_assets]
                    pick = st.selectbox("Choose asset", options=asset_names, key=f"asset_pick_{selected}")
                    asset_obj = next(a for a in matched_assets if a.name == pick)
                else:
                    asset_obj = matched_assets[0]

                preview_asset(asset_obj)

                # Download single asset
                st.download_button(
                    "Download",
                    data=asset_obj.bytes_data,
                    file_name=asset_obj.name,
                    use_container_width=True,
                )

                # Download all assets for this ad code as zip (if more than 1)
                if len(matched_assets) > 1:
                    zbytes = make_assets_zip_for_ad(selected, matched_assets)
                    st.download_button(
                        "Download all matched assets (ZIP)",
                        data=zbytes,
                        file_name=f"assets_{selected}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

else:
    st.info("Upload an Excel (.xlsx) and at least one asset (or a ZIP of assets) to begin.")
