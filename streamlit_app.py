import streamlit as st
import pandas as pd
import re
import zipfile
import io
import json
from datetime import datetime, timezone

# 1. Layout
st.set_page_config(page_title="Ad creative matcher", layout="wide")
st.title("⚡ Ultra-Fast Ad Matcher (XLSX) + Session Save/Restore")

AD_CODE_RE = re.compile(r"\b\d{8}\b")


def normalize_ad_code(value) -> str | None:
    s = "" if value is None else str(value).strip()

    # Prefer exact 8-digit match
    m = AD_CODE_RE.search(s)
    if m:
        return m.group(0)

    # Fallback: digits-only handling
    digits = re.sub(r"\D", "", s)

    # Excel float case: "48725703.0" -> "487257030"
    if len(digits) == 9 and digits.endswith("0") and len(digits[:-1]) == 8:
        return digits[:-1]

    if len(digits) == 8:
        return digits

    # If longer, take first 8 (safer than last-8)
    if len(digits) > 8:
        return digits[:8]

    return None


def extract_ad_code_from_filename(filename: str) -> str | None:
    if not filename:
        return None

    # best: direct 8-digit match anywhere
    m = AD_CODE_RE.search(filename)
    if m:
        return m.group(0)

    # fallback: first long digit run, take first 8
    runs = re.findall(r"\d{8,}", filename)
    if runs:
        return runs[0][:8]

    return None


@st.cache_data
def load_excel_smart(excel_bytes: bytes):
    """
    Reads XLSX bytes.
    Detects header row by scanning for "Ad Code".
    Returns df, ad_code_col, ad_codes(list[str])
    """
    raw = pd.read_excel(io.BytesIO(excel_bytes), engine="openpyxl", header=None)

    header_row_idx = None
    for i in range(min(len(raw), 80)):
        row = raw.iloc[i].astype(str).str.strip().str.lower()
        if row.str.contains(r"\bad\s*code\b", regex=True, na=False).any():
            header_row_idx = i
            break
    if header_row_idx is None:
        header_row_idx = 0

    df = pd.read_excel(io.BytesIO(excel_bytes), engine="openpyxl", header=header_row_idx)
    df.columns = [str(c).strip().lower() for c in df.columns]

    ad_code_col = None
    for c in df.columns:
        if re.search(r"\bad\s*code\b", c):
            ad_code_col = c
            break

    ad_codes = []
    if ad_code_col and ad_code_col in df.columns:
        for v in df[ad_code_col].tolist():
            code = normalize_ad_code(v)
            if code:
                ad_codes.append(code)
    else:
        all_text = "\n".join(df.astype(str).fillna("").values.ravel().tolist())
        ad_codes = AD_CODE_RE.findall(all_text)

    ad_codes = sorted(set(ad_codes))
    return df, ad_code_col, ad_codes


@st.cache_resource
def load_assets(uploaded_files):
    """
    Loads uploaded files/zip(s) into memory.
    Returns list of dicts: {name,data,ext}
    """
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


@st.cache_data
def build_asset_index(all_assets):
    """
    index: {ad_code: [assets...]}
    unparsed: [asset names without code]
    """
    index = {}
    unparsed = []

    for a in all_assets:
        code = extract_ad_code_from_filename(a["name"])
        if not code:
            unparsed.append(a["name"])
            continue
        index.setdefault(code, []).append(a)

    return index, unparsed


def get_rows_for_code(df: pd.DataFrame, ad_code_col: str | None, code: str) -> pd.DataFrame:
    if df is None or not ad_code_col or ad_code_col not in df.columns:
        return pd.DataFrame()

    target = normalize_ad_code(code)
    if not target:
        return pd.DataFrame()

    norm_col = df[ad_code_col].apply(normalize_ad_code)
    return df[norm_col == target]


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


def restore_session_zip(session_zip_file):
    """
    Reads a "session zip" and returns:
      excel_bytes, assets(list)
    Expected structure:
      session/ad_details.xlsx
      session/assets/<files...>
      session/session_meta.json (optional)
    """
    excel_bytes = None
    assets = []

    with zipfile.ZipFile(session_zip_file) as z:
        # Find excel
        # Prefer session/ad_details.xlsx, else first .xlsx found
        names = z.namelist()
        preferred = [n for n in names if n.lower().endswith("session/ad_details.xlsx")]
        if preferred:
            excel_name = preferred[0]
        else:
            xlsx_files = [n for n in names if n.lower().endswith(".xlsx")]
            excel_name = xlsx_files[0] if xlsx_files else None

        if excel_name:
            with z.open(excel_name) as f:
                excel_bytes = f.read()

        # Load assets under session/assets/
        for n in names:
            if n.endswith("/") or "__MACOSX" in n:
                continue
            if n.lower().startswith("session/assets/"):
                with z.open(n) as f:
                    data = f.read()
                # Keep original filename AFTER session/assets/
                original_name = n[len("session/assets/") :]
                if not original_name:
                    continue
                assets.append(
                    {
                        "name": original_name,
                        "data": data,
                        "ext": original_name.split(".")[-1].lower() if "." in original_name else "",
                    }
                )

    return excel_bytes, assets


def build_session_zip_bytes(excel_bytes: bytes, all_assets: list[dict]) -> bytes:
    """
    Builds a single ZIP containing excel + all assets + meta.
    NOTE: This builds in-memory; very large sessions may exceed RAM on Streamlit Cloud.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
        # Excel
        z.writestr("session/ad_details.xlsx", excel_bytes)

        # Assets
        for a in all_assets:
            safe_name = a["name"].lstrip("/").replace("\\", "/")
            z.writestr(f"session/assets/{safe_name}", a["data"])

        # Meta
        meta = {
            "created_utc": datetime.now(timezone.utc).isoformat(),
            "asset_count": len(all_assets),
        }
        z.writestr("session/session_meta.json", json.dumps(meta, indent=2))

    return buf.getvalue()


# Sidebar: upload + restore + reset
with st.sidebar:
    st.header("Upload / Restore")

    restore_zip = st.file_uploader("Restore Session ZIP", type=["zip"])
    st.divider()

    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"], disabled=bool(restore_zip))
    raw_files = st.file_uploader(
        "Upload Assets or ZIP",
        accept_multiple_files=True,
        disabled=bool(restore_zip),
    )

    if st.button("Clear Cache / Reset"):
        st.cache_resource.clear()
        st.cache_data.clear()
        st.session_state.clear()
        st.rerun()


# Decide source: restored session OR manual upload
excel_bytes = None
all_assets = None

if restore_zip:
    excel_bytes, all_assets = restore_session_zip(restore_zip)
    if not excel_bytes:
        st.error("Restore ZIP did not contain an Excel (.xlsx). Expected session/ad_details.xlsx")
    if not all_assets:
        st.warning("Restore ZIP contained no assets under session/assets/")
elif excel_file and raw_files:
    excel_bytes = excel_file.getvalue()
    all_assets = load_assets(raw_files)

# Main
if excel_bytes and all_assets is not None:
    df, ad_code_col, ad_codes = load_excel_smart(excel_bytes)
    asset_index, unparsed = build_asset_index(all_assets)

    st.success(
        f"Loaded {len(all_assets)} assets. Found {len(ad_codes)} Ad Codes in Excel. "
        f"Parsed {len(asset_index)} unique Ad Codes from filenames."
    )

    # SAVE SESSION ZIP BUTTON
    st.subheader("💾 Save this session")
    st.caption("This downloads a ZIP with the Excel + all uploaded assets so someone can restore instantly later.")
    try:
        session_zip_bytes = build_session_zip_bytes(excel_bytes, all_assets)
        st.download_button(
            "⬇️ Download Session ZIP",
            data=session_zip_bytes,
            file_name="ad_matcher_session.zip",
            mime="application/zip",
            key="download_session_zip",
        )
    except Exception as e:
        st.error(
            "Could not build session ZIP (likely too large for in-memory zipping on Streamlit Cloud)."
        )
        st.exception(e)

    if unparsed:
        with st.expander(f"⚠️ {len(unparsed)} assets had no 8-digit ad code in filename"):
            st.write(unparsed[:200])
            if len(unparsed) > 200:
                st.caption("Showing first 200 only.")

    # Match summary
    summary_rows = [{"Ad Code": code, "Matched Assets": len(asset_index.get(code, []))} for code in ad_codes]
    summary_df = (
        pd.DataFrame(summary_rows)
        .sort_values(["Matched Assets", "Ad Code"], ascending=[False, True])
        .reset_index(drop=True)
    )

    st.subheader("✅ Match Summary")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    search = st.text_input("🔍 Quick Search Ad Code", placeholder="Type to filter...")
    filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes

    only_matched = st.checkbox("Show only Ad Codes with matches", value=True)
    if only_matched:
        filtered_codes = [c for c in filtered_codes if c in asset_index]

    for code in filtered_codes:
        matches = asset_index.get(code, [])
        with st.expander(f"✅ Ad {code} — {len(matches)} matched files", expanded=True):
            c1, c2 = st.columns([1, 1.5])

            with c1:
                st.markdown("**Ad Details (copy these):**")
                rows = get_rows_for_code(df, ad_code_col, code)

                if len(rows) > 0:
                    st.dataframe(rows, use_container_width=True, hide_index=True)
                    first = rows.iloc[0].to_dict()
                    pretty = "\n".join(
                        [f"{k}: {v}" for k, v in first.items() if str(v).strip().lower() != "nan"]
                    )
                    st.text_area("Copy-friendly details", pretty, height=240, key=f"copy_{code}")
                else:
                    st.warning("No matching Excel row found for this Ad Code.")
                    st.write("Detected Ad Code column:", ad_code_col)

            with c2:
                st.markdown("**Creative Preview + Download:**")
                if not matches:
                    st.error("No assets matched this Ad Code (from filename parsing).")
                else:
                    for asset in matches:
                        st.caption(f"File: {asset['name']}")
                        preview_asset(asset)
                        st.download_button(
                            "Download",
                            data=asset["data"],
                            file_name=asset["name"].split("/")[-1],
                            key=f"dl_{asset['name']}_{code}",
                        )
                        st.divider()

else:
    st.info("Upload Excel + assets OR restore a previously saved session ZIP.")
