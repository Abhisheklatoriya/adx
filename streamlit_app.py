import streamlit as st
import pandas as pd
import re
import zipfile
import io

st.set_page_config(page_title="Ad creative matcher", layout="wide")
st.title("⚡ Ultra-Fast Ad Matcher (XLSX)")

AD_CODE_RE = re.compile(r"\b\d{8}\b")


def normalize_ad_code(value) -> str | None:
    """
    Returns an 8-digit ad code string if found, else None.
    Handles:
      - 48725703
      - 48725703.0
      - "Ad Code: 48725703"
      - "48,725,703"
    """
    s = "" if value is None else str(value).strip()

    m = AD_CODE_RE.search(s)
    if m:
        return m.group(0)

    digits = re.sub(r"\D", "", s)

    # Excel float case: "48725703.0" => digits "487257030"
    if len(digits) == 9 and digits.endswith("0") and len(digits[:-1]) == 8:
        return digits[:-1]

    if len(digits) == 8:
        return digits

    if len(digits) > 8:
        # If extra digits exist, usually first 8 is the ad code in these sheets
        return digits[:8]

    return None


def extract_ad_code_from_filename(filename: str) -> str | None:
    """
    Extract an 8-digit ad code from filenames like:
      asset_ad_48734339_WAUOgJ.mp4

    Strategy:
      1) Prefer an explicit 8-digit match anywhere in filename.
      2) Fallback: take the first 8 digits from any longer digit run.
    """
    if not filename:
        return None

    m = AD_CODE_RE.search(filename)
    if m:
        return m.group(0)

    # fallback: any long digit run
    runs = re.findall(r"\d{8,}", filename)
    if runs:
        return runs[0][:8]

    return None


@st.cache_data
def load_excel_smart(uploaded_file):
    """
    Loads XLSX and detects header row by scanning for "Ad Code".
    Returns: df (normalized col names), ad_code_col (or None), ad_codes (list[str])
    """
    content = uploaded_file.getvalue()

    raw = pd.read_excel(io.BytesIO(content), engine="openpyxl", header=None)

    header_row_idx = None
    for i in range(min(len(raw), 80)):
        row = raw.iloc[i].astype(str).str.strip().str.lower()
        if row.str.contains(r"\bad\s*code\b", regex=True, na=False).any():
            header_row_idx = i
            break
    if header_row_idx is None:
        header_row_idx = 0

    df = pd.read_excel(io.BytesIO(content), engine="openpyxl", header=header_row_idx)
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
    Builds:
      - index: {ad_code: [assets...]}
      - unparsed: [asset names with no code found]
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


# Sidebar
with st.sidebar:
    st.header("Upload")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    raw_files = st.file_uploader("Upload Assets or ZIP", accept_multiple_files=True)

    if st.button("Clear Cache / Reset"):
        st.cache_resource.clear()
        st.cache_data.clear()
        st.rerun()


if excel_file and raw_files:
    df, ad_code_col, ad_codes = load_excel_smart(excel_file)
    all_assets = load_assets(raw_files)
    asset_index, unparsed = build_asset_index(all_assets)

    st.success(
        f"Loaded {len(all_assets)} asset files. Found {len(ad_codes)} Ad Codes in Excel. "
        f"Parsed {len(asset_index)} unique Ad Codes from filenames."
    )

    if unparsed:
        with st.expander(f"⚠️ {len(unparsed)} assets had no 8-digit code in filename (click to view)"):
            st.write(unparsed[:200])
            if len(unparsed) > 200:
                st.caption("Showing first 200 only.")

    # Match Summary
    summary_rows = []
    for code in ad_codes:
        summary_rows.append({"Ad Code": code, "Matched Assets": len(asset_index.get(code, []))})

    summary_df = (
        pd.DataFrame(summary_rows)
        .sort_values(["Matched Assets", "Ad Code"], ascending=[False, True])
        .reset_index(drop=True)
    )

    st.subheader("✅ Match Summary")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    # Search / filter
    search = st.text_input("🔍 Quick Search Ad Code", placeholder="Type to filter...")
    filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes

    only_matched = st.checkbox("Show only Ad Codes with matches", value=True)
    if only_matched:
        filtered_codes = [c for c in filtered_codes if c in asset_index]

    # Main display
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
                    st.text_area("Copy-friendly details", pretty, height=240)
                else:
                    st.warning("No matching Excel row found for this Ad Code.")
                    st.write("Detected Ad Code column:", ad_code_col)

            with c2:
                st.markdown("**Creative Preview + Download:**")
                if not matches:
                    st.error("No asset filenames parsed to this Ad Code.")
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
    st.info("Please upload your Excel file and creative assets to begin.")
