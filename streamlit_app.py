# streamlit_app.py
import io
import os
import re
import time
import json
import zipfile
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


# -------------------------------
# Page
# -------------------------------
st.set_page_config(page_title="Ultra-Fast Ad Matcher (XLSX)", layout="wide")
APP_TITLE = "⚡ Ultra-Fast Ad Matcher (XLSX)"
SESSION_VERSION = "1.1"

# Example asset filename:
# asset_ad_48734339_WAUOgJ.mp4
ASSET_ADCODE_REGEX = re.compile(r"(?:^|[_\-\s])asset[_\-\s]?ad[_\-\s]?(\d+)", re.IGNORECASE)
FALLBACK_ADCODE_REGEX = re.compile(r"(?:^|[_\-\s])ad[_\-\s]?(\d+)", re.IGNORECASE)

ALLOWED_ASSET_EXTS = {
    ".mp4", ".mov", ".m4v", ".webm",
    ".jpg", ".jpeg", ".png", ".gif", ".webp",
    ".pdf",
    ".zip",
    ".html", ".htm",
    ".mp3", ".wav",
}

# Streamlit is deprecating use_container_width; new API uses width="stretch"
STRETCH = "stretch"


# -------------------------------
# Helpers
# -------------------------------
def _safe_filename(name: str) -> str:
    name = os.path.basename(str(name))
    name = re.sub(r"[^a-zA-Z0-9._\- ]+", "_", name).strip()
    return name[:200] if name else "file"


def normalize_str(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())


def extract_ad_code_from_asset_filename(filename: str) -> Optional[str]:
    base = os.path.basename(str(filename))
    m = ASSET_ADCODE_REGEX.search(base)
    if m:
        return m.group(1)
    m2 = FALLBACK_ADCODE_REGEX.search(base)
    if m2:
        return m2.group(1)
    return None


@st.cache_data
def load_excel(file_like) -> Tuple[pd.DataFrame, Optional[str]]:
    """
    Loads Excel and tries to auto-detect an ad-code-like column.
    If not found, returns (df, None) so UI can ask user to pick.
    """
    df = pd.read_excel(file_like, engine="openpyxl")
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


def load_assets(uploaded_files: List) -> List[AssetItem]:
    """
    Accepts files and/or ZIP(s). If ZIP, extracts all files inside.
    Stores bytes in memory for preview + download.
    """
    assets: List[AssetItem] = []

    for uf in uploaded_files:
        fname = _safe_filename(uf.name)
        ext = os.path.splitext(fname)[1].lower()
        raw = _read_uploaded_file_bytes(uf)

        if ext == ".zip":
            with zipfile.ZipFile(io.BytesIO(raw), "r") as zf:
                for info in zf.infolist():
                    if info.is_dir():
                        continue
                    inner_name = _safe_filename(info.filename)
                    inner_ext = os.path.splitext(inner_name)[1].lower()
                    data = zf.read(info.filename)
                    ad_code = extract_ad_code_from_asset_filename(inner_name)
                    assets.append(
                        AssetItem(name=inner_name, ext=inner_ext, ad_code=ad_code, bytes_data=data)
                    )
        else:
            ad_code = extract_ad_code_from_asset_filename(fname)
            assets.append(AssetItem(name=fname, ext=ext, ad_code=ad_code, bytes_data=raw))

    return assets


def build_matches(df: pd.DataFrame, ad_code_col: str, assets: List[AssetItem]) -> Tuple[pd.DataFrame, Dict[str, List[AssetItem]]]:
    """
    Matching logic:
    - Excel ad code might be "48726907D" -> we extract the first digit-run "48726907"
    - Asset ad code extracted from filename "asset_ad_48726907_xxx.mp4" -> "48726907"
    """
    df = df.copy()

    def excel_to_key(x) -> Optional[str]:
        if pd.isna(x):
            return None
        s = str(x).strip()
        m = re.search(r"(\d+)", s)
        return m.group(1) if m else None

    df["__ad_code_key__"] = df[ad_code_col].apply(excel_to_key)

    match_map: Dict[str, List[AssetItem]] = {}
    for a in assets:
        if not a.ad_code:
            continue
        match_map.setdefault(a.ad_code, []).append(a)

    return df, match_map


def preview_asset(asset: AssetItem):
    st.caption(f"File: {asset.name}")
    ext = asset.ext.lower()

    if ext in [".mp4", ".mov", ".m4v", ".webm"]:
        st.video(asset.bytes_data)
    elif ext in [".mp3", ".wav"]:
        st.audio(asset.bytes_data)
    elif ext in [".jpg", ".jpeg", ".png", ".gif", ".webp"]:
        st.image(asset.bytes_data, width=STRETCH)
    elif ext == ".pdf":
        st.info("PDF preview isn't native in Streamlit; use the download button.")
    elif ext in [".html", ".htm"]:
        try:
            html = asset.bytes_data.decode("utf-8", errors="ignore")
            components.html(html, height=500, scrolling=True)
        except Exception:
            st.info("Couldn't render HTML preview; use the download button.")
    else:
        st.info("No preview for this file type. Use the download button.")


def copy_block(text: str, key: str):
    """
    Textarea + Copy button (clipboard) using JS.
    """
    st.text_area("Copy-friendly details", value=text, height=260, key=key)

    if st.button("Copy", key=f"{key}_copybtn"):
        try:
            # Escape for JS template literal
            escaped = text.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
            components.html(
                f"""
                <script>
                  const text = `{escaped}`;
                  navigator.clipboard.writeText(text).then(() => {{
                    // no-op
                  }});
                </script>
                """,
                height=0,
            )
            st.success("Copied to clipboard.")
        except Exception as e:
            st.error(f"Copy failed: {e}")


def make_session_zip_bytes(excel_name: str, excel_bytes: bytes, assets: List[AssetItem], meta: dict) -> bytes:
    """
    Creates a 'session bundle' ZIP containing:
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
    Loads a previously saved session ZIP.
    Returns: (excel_bytes, excel_name, assets, meta)
    """
    excel_bytes = None
    excel_name = None
    assets: List[AssetItem] = []
    meta = {}

    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
        if "session/meta.json" in zf.namelist():
            meta = json.loads(zf.read("session/meta.json").decode("utf-8", errors="ignore"))

        excel_candidates = [n for n in zf.namelist() if n.startswith("inputs/excel/") and not n.endswith("/")]
        if excel_candidates:
            excel_name = os.path.basename(excel_candidates[0])
            excel_bytes = zf.read(excel_candidates[0])

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
# UI
# -------------------------------
st.title(APP_TITLE)

with st.sidebar:
    st.header("Upload")

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
                st.success("Session loaded. Go to the main area.")
        except Exception as e:
            st.error(f"Failed to load session ZIP: {e}")

    st.divider()

    st.subheader("Upload Excel (.xlsx)")
    excel_file = st.file_uploader("Excel file", type=["xlsx"], key="excel_uploader")

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
        st.success("Cleared. Refresh if needed.")


# -------------------------------
# Acquire inputs
# -------------------------------
excel_bytes = None
excel_name = None
assets: List[AssetItem] = []

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
# Main
# -------------------------------
if excel_bytes and assets:
    df, detected_col = load_excel(io.BytesIO(excel_bytes))

    ad_code_col = detected_col
    if ad_code_col is None:
        st.warning("Excel load: couldn't auto-detect the Ad Code column. Please select it.")
        ad_code_col = st.selectbox("Select the Ad Code column", options=list(df.columns), key="ad_code_picker")

    df_clean, match_map = build_matches(df, ad_code_col, assets)

    available_adcodes = sorted([k for k in df_clean["__ad_code_key__"].dropna().unique().tolist() if k])
    st.success(f"Loaded {len(assets)} assets. Found {len(available_adcodes)} Ad Codes in Excel.")

    # Session download
    st.subheader("Session")
    meta = {
        "version": SESSION_VERSION,
        "created_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "excel_name": excel_name,
        "asset_count": len(assets),
        "notes": "Upload this ZIP later to restore the session without re-uploading everything.",
    }
    session_zip_bytes = make_session_zip_bytes(excel_name or "input.xlsx", excel_bytes, assets, meta)
    st.download_button(
        "Download session as ZIP",
        data=session_zip_bytes,
        file_name=f"ad_match_session_{int(time.time())}.zip",
        mime="application/zip",
        width=STRETCH,
    )

    st.divider()

    st.markdown("🔎 **Quick Search Ad Code**")
    q = st.text_input("Type to filter…", value="", key="search_box").strip()

    browse = df_clean.copy()
    browse["matched_assets"] = browse["__ad_code_key__"].apply(lambda k: len(match_map.get(k, [])) if k else 0)

    if q:
        browse = browse[browse["__ad_code_key__"].astype(str).str.contains(q, na=False)]

    show_cols = []
    preferred = [ad_code_col, "media", "first run date", "last run date", "brand", "headline", "description"]
    lower_to_real = {c.lower(): c for c in browse.columns}
    for p in preferred:
        if p is None:
            continue
        if p in browse.columns:
            show_cols.append(p)
        elif p.lower() in lower_to_real:
            show_cols.append(lower_to_real[p.lower()])

    if "__ad_code_key__" not in show_cols:
        show_cols.insert(0, "__ad_code_key__")
    if "matched_assets" not in show_cols:
        show_cols.insert(1, "matched_assets")

    st.dataframe(
        browse[show_cols].rename(columns={"__ad_code_key__": "ad code key"}),
        width=STRETCH,
        hide_index=True,
    )

    st.divider()

    left, right = st.columns([0.55, 0.45], gap="large")

    with left:
        st.subheader("Pick an Ad Code to preview")
        if not available_adcodes:
            st.error("No usable Ad Codes found in the selected column (no digits detected).")
            st.stop()

        selected = st.selectbox(
            "Ad Code (numeric key)",
            options=available_adcodes,
            index=0,
            key="adcode_select",
        )

        rows = df_clean[df_clean["__ad_code_key__"] == selected].copy()

        st.markdown("**Matched Excel rows**")
        st.dataframe(rows.drop(columns=["__ad_code_key__"]), width=STRETCH, hide_index=True)

        r0 = rows.iloc[0].to_dict()
        lines = [
            f"ad code: {r0.get(ad_code_col)}",
            f"ad code key: {selected}",
        ]
        for col in df_clean.columns:
            if col.startswith("__") or col == ad_code_col:
                continue
            val = r0.get(col)
            if pd.isna(val):
                continue
            s = str(val)
            if len(s) > 250:
                s = s[:250] + "…"
            lines.append(f"{col}: {s}")

        detail_text = "\n".join(lines)
        copy_block(detail_text, key=f"copy_{selected}")

    with right:
        st.subheader("Matched asset preview")
        matched_assets = match_map.get(selected, [])
        if not matched_assets:
            st.warning("No assets matched this Ad Code key.")
        else:
            if len(matched_assets) > 1:
                asset_names = [a.name for a in matched_assets]
                pick = st.selectbox("Choose asset", options=asset_names, key=f"asset_pick_{selected}")
                asset_obj = next(a for a in matched_assets if a.name == pick)
            else:
                asset_obj = matched_assets[0]

            preview_asset(asset_obj)

            st.download_button(
                "Download asset",
                data=asset_obj.bytes_data,
                file_name=asset_obj.name,
                width=STRETCH,
            )

            if len(matched_assets) > 1:
                zbytes = make_assets_zip_for_ad(selected, matched_assets)
                st.download_button(
                    "Download all matched assets (ZIP)",
                    data=zbytes,
                    file_name=f"assets_{selected}.zip",
                    mime="application/zip",
                    width=STRETCH,
                )

else:
    st.info("Upload an Excel (.xlsx) and at least one asset (or a ZIP of assets) to begin.")
