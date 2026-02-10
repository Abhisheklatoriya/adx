import streamlit as st
import pandas as pd
import re
import zipfile
import hashlib

# -------------------------------
# App Config
# -------------------------------
st.set_page_config(page_title="Ad Creative Matcher", layout="wide")
st.title("⚡ Ultra-Fast Ad Matcher (Excel + Live Loading)")

# -------------------------------
# Helpers
# -------------------------------
def _file_sig(uploaded_file) -> str:
    """
    Stable signature for 'already processed' tracking.
    Uses name + size + a small hash of the first ~1MB (fast enough, avoids full read).
    """
    try:
        data = uploaded_file.getvalue()
        head = data[:1024 * 1024]  # 1MB
        h = hashlib.md5(head).hexdigest()
        return f"{uploaded_file.name}|{len(data)}|{h}"
    except Exception:
        # Fallback
        return f"{uploaded_file.name}"

@st.cache_data
def load_excel(file):
    df = pd.read_excel(file)
    df.columns = [c.strip().lower() for c in df.columns]

    ad_code_col = next((c for c in df.columns if "ad" in c and "code" in c), None)
    if not ad_code_col:
        raise ValueError("No Ad Code column found (expected something like 'Ad Code').")

    df[ad_code_col] = df[ad_code_col].astype(str).str.strip()
    df = df[df[ad_code_col].str.match(r"^\d{8}$")]

    # optional: ensure unique per ad code (keep first row)
    # df = df.drop_duplicates(subset=[ad_code_col], keep="first")

    return df, ad_code_col

def add_asset(name: str, data: bytes):
    ext = name.split(".")[-1].lower() if "." in name else ""
    st.session_state.assets.append({"name": name, "data": data, "ext": ext})

def process_new_uploads(uploaded_files):
    """
    Incrementally process only new uploads, updating progress live.
    """
    if not uploaded_files:
        return

    progress = st.session_state.ui_progress
    status = st.session_state.ui_status

    # Count how many total "steps" we *roughly* have:
    # - Non-zip files: 1 step each
    # - Zip files: number of entries (we’ll compute per zip)
    total_steps = 0
    zip_entry_counts = {}

    for f in uploaded_files:
        sig = _file_sig(f)
        if sig in st.session_state.processed_sigs:
            continue

        if f.name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(f) as z:
                    entries = [i for i in z.infolist() if (not i.is_dir()) and "__MACOSX" not in i.filename]
                    zip_entry_counts[sig] = len(entries)
                    total_steps += max(len(entries), 1)
            except Exception:
                total_steps += 1
        else:
            total_steps += 1

    if total_steps == 0:
        return

    done_steps = 0

    for f in uploaded_files:
        sig = _file_sig(f)
        if sig in st.session_state.processed_sigs:
            continue

        status.write(f"Processing: **{f.name}**")
        if f.name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(f) as z:
                    entries = [i for i in z.infolist() if (not i.is_dir()) and "__MACOSX" not in i.filename]
                    if not entries:
                        # empty zip
                        done_steps += 1
                        progress.progress(min(done_steps / total_steps, 1.0))
                    else:
                        for info in entries:
                            with z.open(info) as fh:
                                add_asset(info.filename, fh.read())

                            done_steps += 1
                            progress.progress(min(done_steps / total_steps, 1.0))
            except Exception as e:
                st.warning(f"Could not read ZIP {f.name}: {e}")
                done_steps += 1
                progress.progress(min(done_steps / total_steps, 1.0))
        else:
            # Normal file
            try:
                add_asset(f.name, f.getvalue())
            except Exception as e:
                st.warning(f"Could not read file {f.name}: {e}")

            done_steps += 1
            progress.progress(min(done_steps / total_steps, 1.0))

        st.session_state.processed_sigs.add(sig)

    status.write("✅ Processing complete for currently uploaded files.")

# -------------------------------
# Session State
# -------------------------------
if "assets" not in st.session_state:
    st.session_state.assets = []  # list of {"name","data","ext"}
if "processed_sigs" not in st.session_state:
    st.session_state.processed_sigs = set()

# UI placeholders for live updates
if "ui_progress" not in st.session_state:
    st.session_state.ui_progress = st.empty()
if "ui_status" not in st.session_state:
    st.session_state.ui_status = st.empty()

# -------------------------------
# Sidebar
# -------------------------------
with st.sidebar:
    st.header("Upload")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    raw_files = st.file_uploader("Upload Assets or ZIPs", accept_multiple_files=True)

    if st.button("Reset session"):
        st.session_state.assets = []
        st.session_state.processed_sigs = set()
        st.cache_data.clear()
        st.rerun()

# -------------------------------
# Main
# -------------------------------
if not excel_file:
    st.info("Upload your Excel ad list to begin.")
    st.stop()

try:
    df, ad_code_col = load_excel(excel_file)
except Exception as e:
    st.error(f"Excel load error: {e}")
    st.stop()

# Live processing UI
st.markdown("### 📦 Live asset loading")
progress_bar = st.progress(0.0)
st.session_state.ui_progress = progress_bar
status_box = st.empty()
st.session_state.ui_status = status_box

# Process newly uploaded files incrementally
if raw_files:
    process_new_uploads(raw_files)
else:
    status_box.write("Upload files/ZIPs to start matching.")

# Summary
st.success(f"Ads in Excel: **{len(df)}** · Assets loaded so far: **{len(st.session_state.assets)}**")

# Search and match as assets accumulate
search = st.text_input("🔍 Search Ad Code", placeholder="Type partial or full 8-digit code...")
ad_codes = df[ad_code_col].tolist()
filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes

# Quick toggle: show only matched codes
only_matched = st.checkbox("Show only ad codes with matches", value=True)

assets = st.session_state.assets

# Build matches live
for ad_code in filtered_codes:
    matches = [a for a in assets if ad_code in a["name"]]

    if only_matched and not matches:
        continue

    with st.expander(f"{'✅' if matches else '⚪️'} Ad {ad_code} — {len(matches)} files", expanded=bool(matches)):
        c1, c2 = st.columns([1, 1.6])

        with c1:
            st.markdown("**Ad Specs (Excel row)**")
            row = df[df[ad_code_col] == ad_code]
            st.dataframe(row, use_container_width=True)

        with c2:
            if not matches:
                st.caption("No matching files yet. Keep uploading…")
            else:
                for asset in matches:
                    st.caption(asset["name"])

                    ext = asset["ext"]
                    if ext in ["mp4", "mov", "webm"]:
                        st.video(asset["data"])
                    elif ext in ["mp3", "wav"]:
                        st.audio(asset["data"])
                    elif ext in ["jpg", "jpeg", "png", "gif"]:
                        st.image(asset["data"])

                    st.download_button(
                        "Download",
                        data=asset["data"],
                        file_name=asset["name"].split("/")[-1],  # safer filename
                        key=f"{ad_code}_{asset['name']}"
                    )
                    st.divider()
