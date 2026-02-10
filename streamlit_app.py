import streamlit as st
import pandas as pd
import zipfile
import re
import os
import shutil
import tempfile
from pathlib import Path

# -------------------------------
# App Config
# -------------------------------
st.set_page_config(page_title="Ad Creative Matcher", layout="wide")
st.title("⚡ Ad Matcher (Excel + Live Processing + Disk Storage)")

# -------------------------------
# Excel loader
# -------------------------------
@st.cache_data
def load_excel(file):
    df = pd.read_excel(file)
    df.columns = [c.strip().lower() for c in df.columns]

    ad_code_col = next((c for c in df.columns if "ad" in c and "code" in c), None)
    if not ad_code_col:
        raise ValueError("Could not find an Ad Code column (expected something like 'Ad Code').")

    df[ad_code_col] = df[ad_code_col].astype(str).str.strip()
    df = df[df[ad_code_col].str.match(r"^\d{8}$")]

    codes = df[ad_code_col].tolist()
    codes_set = set(codes)

    return df, ad_code_col, codes, codes_set

# -------------------------------
# Disk workspace (per session)
# -------------------------------
def ensure_workspace():
    if "workspace_dir" not in st.session_state:
        st.session_state.workspace_dir = tempfile.mkdtemp(prefix="admatcher_")
    return Path(st.session_state.workspace_dir)

def get_output_dir():
    base = ensure_workspace()
    out = base / "output"
    out.mkdir(parents=True, exist_ok=True)
    return out

# -------------------------------
# Helpers
# -------------------------------
CODE_RE = re.compile(r"\b(\d{8})\b")

def extract_code_from_name(name: str, valid_codes_set: set[str]) -> str | None:
    """
    Find an 8-digit code in filename. Return it if it's one of the valid codes from Excel.
    """
    # Try all 8-digit numbers found in the string
    for m in CODE_RE.findall(name):
        if m in valid_codes_set:
            return m
    return None

def safe_filename(p: str) -> str:
    # keep nested folders but avoid weird absolute paths
    p = p.replace("\\", "/")
    p = p.lstrip("/").lstrip("./")
    return p

def write_bytes_to_disk(dest_path: Path, data: bytes):
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    with open(dest_path, "wb") as f:
        f.write(data)

def copy_file_to_disk(uploaded_file, dest_path: Path):
    """
    Streamlit UploadedFile is file-like. We can stream copy to disk without getvalue().
    """
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    uploaded_file.seek(0)
    with open(dest_path, "wb") as out:
        shutil.copyfileobj(uploaded_file, out, length=1024 * 1024)  # 1MB chunks

# -------------------------------
# Incremental processing
# -------------------------------
def process_new_uploads(uploaded_files, valid_codes_set: set[str]):
    """
    Process only new uploads; write matched assets to disk immediately.
    """
    if "processed_names" not in st.session_state:
        st.session_state.processed_names = set()

    out_dir = get_output_dir()

    progress = st.progress(0.0)
    status = st.empty()

    # Rough total steps: count files + zip entries
    steps = 0
    zip_entries_map = {}

    for uf in uploaded_files:
        key = f"{uf.name}|{getattr(uf, 'size', 'na')}"
        if key in st.session_state.processed_names:
            continue

        if uf.name.lower().endswith(".zip"):
            try:
                uf.seek(0)
                with zipfile.ZipFile(uf) as z:
                    entries = [
                        i for i in z.infolist()
                        if (not i.is_dir()) and "__MACOSX" not in i.filename
                    ]
                    zip_entries_map[key] = len(entries)
                    steps += max(len(entries), 1)
            except Exception:
                steps += 1
        else:
            steps += 1

    if steps == 0:
        status.write("No new files to process.")
        return

    done = 0

    for uf in uploaded_files:
        key = f"{uf.name}|{getattr(uf, 'size', 'na')}"
        if key in st.session_state.processed_names:
            continue

        status.write(f"Processing **{uf.name}** …")

        # ZIP: stream entries one-by-one
        if uf.name.lower().endswith(".zip"):
            try:
                uf.seek(0)
                with zipfile.ZipFile(uf) as z:
                    entries = [
                        i for i in z.infolist()
                        if (not i.is_dir()) and "__MACOSX" not in i.filename
                    ]

                    if not entries:
                        done += 1
                        progress.progress(min(done / steps, 1.0))
                    else:
                        for info in entries:
                            fname = safe_filename(info.filename)

                            code = extract_code_from_name(fname, valid_codes_set)
                            if code:
                                dest = out_dir / code / fname
                                with z.open(info) as src:
                                    dest.parent.mkdir(parents=True, exist_ok=True)
                                    with open(dest, "wb") as out:
                                        shutil.copyfileobj(src, out, length=1024 * 1024)

                            done += 1
                            progress.progress(min(done / steps, 1.0))
            except Exception as e:
                st.warning(f"Could not read ZIP {uf.name}: {e}")
                done += 1
                progress.progress(min(done / steps, 1.0))

        # Normal file: copy it if it matches a code
        else:
            fname = safe_filename(uf.name)
            code = extract_code_from_name(fname, valid_codes_set)
            if code:
                dest = out_dir / code / fname
                try:
                    copy_file_to_disk(uf, dest)
                except Exception as e:
                    st.warning(f"Could not save file {uf.name}: {e}")

            done += 1
            progress.progress(min(done / steps, 1.0))

        st.session_state.processed_names.add(key)

    status.write("✅ Done processing currently uploaded files.")

def list_matches_for_code(code: str):
    out_dir = get_output_dir() / code
    if not out_dir.exists():
        return []
    # Return files (recursive)
    return sorted([p for p in out_dir.rglob("*") if p.is_file()])

def ext_of(path: Path) -> str:
    return path.suffix.lower().lstrip(".")

def read_small_preview(path: Path, max_mb=25):
    """
    For previews, don’t load huge files into memory.
    Only read up to max_mb.
    """
    max_bytes = max_mb * 1024 * 1024
    size = path.stat().st_size
    if size > max_bytes:
        return None  # too big to preview safely
    return path.read_bytes()

# -------------------------------
# Sidebar
# -------------------------------
with st.sidebar:
    st.header("Upload")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    raw_files = st.file_uploader("Upload Assets or ZIPs", accept_multiple_files=True)

    if st.button("Reset session"):
        # cleanup workspace
        if "workspace_dir" in st.session_state:
            try:
                shutil.rmtree(st.session_state.workspace_dir, ignore_errors=True)
            except Exception:
                pass
        for k in ["workspace_dir", "processed_names"]:
            if k in st.session_state:
                del st.session_state[k]
        st.cache_data.clear()
        st.rerun()

# -------------------------------
# Main
# -------------------------------
if not excel_file:
    st.info("Upload your Excel ad list to begin.")
    st.stop()

try:
    df, ad_code_col, ad_codes, ad_codes_set = load_excel(excel_file)
except Exception as e:
    st.error(f"Excel load error: {e}")
    st.stop()

st.success(f"Ads in Excel: **{len(df)}**")

if raw_files:
    process_new_uploads(raw_files, ad_codes_set)
else:
    st.info("Upload creatives/ZIPs — matches will appear as soon as files are processed.")

# UI: Search + matches
search = st.text_input("🔍 Search Ad Code", placeholder="Type partial or full 8-digit code…")
filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes

only_matched = st.checkbox("Show only ad codes with matches", value=True)

for code in filtered_codes:
    files = list_matches_for_code(code)

    if only_matched and not files:
        continue

    with st.expander(f"{'✅' if files else '⚪️'} Ad {code} — {len(files)} files", expanded=bool(files)):
        c1, c2 = st.columns([1, 1.6])

        with c1:
            st.markdown("**Ad Specs (Excel row)**")
            st.dataframe(df[df[ad_code_col] == code], use_container_width=True)

        with c2:
            if not files:
                st.caption("No matching files yet. Upload more ZIPs/files…")
            else:
                for p in files:
                    st.caption(str(p.relative_to(get_output_dir() / code)))

                    ext = ext_of(p)

                    # Preview safely (only if small)
                    preview_bytes = read_small_preview(p, max_mb=25)

                    if preview_bytes is None:
                        st.info("Preview skipped (file is large). Download to view.")
                    else:
                        if ext in ["mp4", "mov", "webm"]:
                            st.video(preview_bytes)
                        elif ext in ["mp3", "wav"]:
                            st.audio(preview_bytes)
                        elif ext in ["jpg", "jpeg", "png", "gif"]:
                            st.image(preview_bytes)

                    # Always allow download (streamlit will read bytes for the button)
                    st.download_button(
                        "Download",
                        data=p.read_bytes() if p.stat().st_size <= 200 * 1024 * 1024 else open(p, "rb"),
                        file_name=p.name,
                        key=f"dl_{code}_{p.as_posix()}"
                    )
                    st.divider()
