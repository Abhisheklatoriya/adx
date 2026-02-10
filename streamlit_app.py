import re
import zipfile
import shutil
import tempfile
from pathlib import Path
from typing import Optional, Set, List, Dict, Tuple

import streamlit as st
import pandas as pd


# -------------------------------
# Config
# -------------------------------
st.set_page_config(page_title="Ad Creative Matcher", layout="wide")
st.title("⚡ Ad Creative Matcher (Excel + Live Preview While Processing)")

CODE_RE = re.compile(r"\b(\d{8})\b")
SKIP_ZIP_PATH_PARTS = ("__MACOSX",)


# -------------------------------
# Workspace
# -------------------------------
def ws() -> Path:
    if "workspace" not in st.session_state:
        st.session_state.workspace = tempfile.mkdtemp(prefix="admatcher_")
    return Path(st.session_state.workspace)

def dir_incoming() -> Path:
    p = ws() / "incoming"
    p.mkdir(parents=True, exist_ok=True)
    return p

def dir_output() -> Path:
    p = ws() / "output"
    p.mkdir(parents=True, exist_ok=True)
    return p

def hard_reset():
    if "workspace" in st.session_state:
        shutil.rmtree(st.session_state.workspace, ignore_errors=True)
        del st.session_state["workspace"]
    for k in [
        "processed_upload_keys",
        "queue_files",
        "zip_state",
        "index",
        "processing",
        "selected_code",
    ]:
        if k in st.session_state:
            del st.session_state[k]
    st.cache_data.clear()


# -------------------------------
# Excel loader (auto header row)
# -------------------------------
@st.cache_data
def load_excel_auto_header(excel_file) -> Tuple[pd.DataFrame, str, List[str], Set[str]]:
    raw = pd.read_excel(excel_file, header=None, engine="openpyxl")

    header_row = None
    scan_rows = min(80, len(raw))
    for i in range(scan_rows):
        row_vals = raw.iloc[i].astype(str).str.strip().str.lower()
        if row_vals.str.contains(r"\bad\s*code\b", regex=True).any():
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not find a header row containing 'Ad Code' in the first 80 rows.")

    df = pd.read_excel(excel_file, header=header_row, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    ad_code_col = None
    for c in df.columns:
        cl = c.lower().strip()
        if "ad" in cl and "code" in cl:
            ad_code_col = c
            break

    if not ad_code_col:
        raise ValueError(f"Header found but no Ad Code column. Columns: {list(df.columns)}")

    df[ad_code_col] = df[ad_code_col].astype(str).str.strip()
    df = df[df[ad_code_col].str.match(r"^\d{8}$")]

    ad_codes = df[ad_code_col].tolist()
    return df, ad_code_col, ad_codes, set(ad_codes)


# -------------------------------
# File helpers
# -------------------------------
def safe_relpath(p: str) -> str:
    p = p.replace("\\", "/")
    p = p.lstrip("/").lstrip("./")
    return p

def extract_code_from_name(name: str, valid: Set[str]) -> Optional[str]:
    for m in CODE_RE.findall(name):
        if m in valid:
            return m
    return None

def save_uploaded_to_disk(uploaded_file, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    uploaded_file.seek(0)
    with open(dest, "wb") as out:
        shutil.copyfileobj(uploaded_file, out, length=1024 * 1024)  # 1MB chunks

def is_zip(path: Path) -> bool:
    return path.suffix.lower() == ".zip"

def ext_of(p: Path) -> str:
    return p.suffix.lower().lstrip(".")

def preview_bytes(p: Path, max_mb: int = 25) -> Optional[bytes]:
    if p.stat().st_size > max_mb * 1024 * 1024:
        return None
    return p.read_bytes()

def download_bytes_safe(p: Path, max_mb: int = 200) -> Optional[bytes]:
    # Avoid reading crazy-big files into RAM for download buttons
    if p.stat().st_size > max_mb * 1024 * 1024:
        return None
    return p.read_bytes()


# -------------------------------
# State init
# -------------------------------
def ensure_state():
    if "processed_upload_keys" not in st.session_state:
        st.session_state.processed_upload_keys = set()  # uploads saved to disk
    if "queue_files" not in st.session_state:
        st.session_state.queue_files = []  # list of incoming file paths to process
    if "zip_state" not in st.session_state:
        # active zip processing state:
        # { "zip_path": str, "names": [str...], "i": int }
        st.session_state.zip_state = None
    if "index" not in st.session_state:
        # ad_code -> list of output file paths
        st.session_state.index = {}
    if "processing" not in st.session_state:
        st.session_state.processing = False


# -------------------------------
# Build queue from uploads (disk save)
# -------------------------------
def ingest_uploads_to_queue(uploaded_files: List, ) -> None:
    """
    Save newly uploaded files to disk (incoming/), and push their paths into the processing queue.
    """
    inc = dir_incoming()
    for uf in uploaded_files:
        key = f"{uf.name}|{getattr(uf, 'size', 'na')}"
        if key in st.session_state.processed_upload_keys:
            continue

        # Save upload to disk
        dest = inc / safe_relpath(uf.name)
        save_uploaded_to_disk(uf, dest)

        # Queue the saved file
        st.session_state.queue_files.append(dest.as_posix())
        st.session_state.processed_upload_keys.add(key)


# -------------------------------
# Incremental processor (THIS is what enables live preview)
# -------------------------------
def process_some(valid_codes: Set[str], chunk: int = 200) -> int:
    """
    Process up to `chunk` items (zip entries or files) per rerun.
    Writes matched files to output/<code>/..., updates in-memory index immediately.
    Returns how many items were processed this call.
    """
    out_root = dir_output()
    processed = 0

    # If we're currently inside a zip, continue where we left off
    zs = st.session_state.zip_state
    if zs is not None:
        zip_path = Path(zs["zip_path"])
        names = zs["names"]
        i = zs["i"]

        with zipfile.ZipFile(zip_path) as z:
            while i < len(names) and processed < chunk:
                name = names[i]
                i += 1

                # match code from filename
                code = extract_code_from_name(name, valid_codes)
                if code:
                    dest = out_root / code / safe_relpath(name)
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    with z.open(name) as src, open(dest, "wb") as out:
                        shutil.copyfileobj(src, out, length=1024 * 1024)

                    st.session_state.index.setdefault(code, []).append(dest.as_posix())

                processed += 1

        # save progress
        if i >= len(names):
            st.session_state.zip_state = None
        else:
            st.session_state.zip_state = {"zip_path": zs["zip_path"], "names": names, "i": i}

        return processed

    # Otherwise, pop next file from queue
    while st.session_state.queue_files and processed < chunk:
        p = Path(st.session_state.queue_files.pop(0))

        if is_zip(p):
            # Start zip state (but do not process everything in one go)
            with zipfile.ZipFile(p) as z:
                names = []
                for info in z.infolist():
                    if info.is_dir():
                        continue
                    if any(part in info.filename for part in SKIP_ZIP_PATH_PARTS):
                        continue
                    names.append(info.filename)

            st.session_state.zip_state = {"zip_path": p.as_posix(), "names": names, "i": 0}
            # immediately continue processing zip entries in this same call
            return processed + process_some(valid_codes, chunk=(chunk - processed))

        else:
            # Normal file: match on name and copy if code found
            code = extract_code_from_name(p.name, valid_codes)
            if code:
                dest = out_root / code / safe_relpath(p.name)
                dest.parent.mkdir(parents=True, exist_ok=True)
                with open(p, "rb") as src, open(dest, "wb") as out:
                    shutil.copyfileobj(src, out, length=1024 * 1024)

                st.session_state.index.setdefault(code, []).append(dest.as_posix())

            processed += 1

    return processed


# -------------------------------
# UI: Two-pane preview
# -------------------------------
def render_two_pane(df: pd.DataFrame, ad_code_col: str, code: str):
    c1, c2 = st.columns([1, 1.6], vertical_alignment="top")

    with c1:
        st.markdown("### Ad Info")
        st.dataframe(df[df[ad_code_col] == code], use_container_width=True)

    with c2:
        st.markdown("### Matched Creatives")
        files = st.session_state.index.get(code, [])
        if not files:
            st.info("No creatives matched yet for this ad code (keep uploading / processing).")
            return

        # Show newest first
        for fp in reversed(files):
            p = Path(fp)
            st.caption(str(p.relative_to(dir_output() / code)))

            ext = ext_of(p)
            pv = preview_bytes(p, max_mb=25)

            if pv is None:
                st.warning("Preview skipped (file too large). You can still download if under 200MB.")
            else:
                if ext in ["mp4", "mov", "webm"]:
                    st.video(pv)
                elif ext in ["mp3", "wav"]:
                    st.audio(pv)
                elif ext in ["jpg", "jpeg", "png", "gif", "webp"]:
                    st.image(pv)

            dl = download_bytes_safe(p, max_mb=200)
            if dl is None:
                st.error("Download disabled for very large file (>200MB). If you need this, tell me and I’ll add a zip-per-ad export.")
            else:
                st.download_button(
                    "Download",
                    data=dl,
                    file_name=p.name,
                    key=f"dl_{code}_{p.as_posix()}",
                )

            st.divider()


# -------------------------------
# Sidebar
# -------------------------------
ensure_state()

with st.sidebar:
    st.header("Upload")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    uploads = st.file_uploader("Upload Assets or ZIPs", accept_multiple_files=True)

    st.divider()
    st.session_state.processing = st.toggle("Auto-process (live)", value=st.session_state.processing)

    chunk = st.slider("Processing chunk per refresh", 50, 1000, 200, 50)

    if st.button("Reset / Clear Session"):
        hard_reset()
        st.rerun()


# -------------------------------
# Main: Excel
# -------------------------------
if not excel_file:
    st.info("Upload the Excel ad list to begin.")
    st.stop()

try:
    df, ad_code_col, ad_codes, ad_codes_set = load_excel_auto_header(excel_file)
except Exception as e:
    st.error(f"Excel load error: {e}")
    st.stop()

st.success(f"Loaded **{len(df)}** ads. Detected Ad Code column: **{ad_code_col}**")

# Ingest uploads into queue (disk save)
if uploads:
    ingest_uploads_to_queue(uploads)

# -------------------------------
# Processing loop (incremental across reruns)
# -------------------------------
queue_len = len(st.session_state.queue_files)
zip_state = st.session_state.zip_state
active_zip = zip_state["zip_path"] if zip_state else None

status_row = st.container()
with status_row:
    colA, colB, colC = st.columns(3)
    colA.metric("Queued files", queue_len)
    colB.metric("Active ZIP", "Yes" if active_zip else "No")
    colC.metric("Matched ad codes", len(st.session_state.index))

# If auto-processing is enabled, process a chunk and auto-refresh
if st.session_state.processing and (queue_len > 0 or zip_state is not None):
    processed_now = process_some(ad_codes_set, chunk=chunk)
    st.info(f"Processing… just handled {processed_now} items this refresh.")
    # Auto refresh quickly while work remains
    st.autorefresh(interval=700, key="auto_refresh_processing")

elif (queue_len > 0 or zip_state is not None):
    st.warning("Uploads queued. Turn on **Auto-process (live)** to start processing, or add a button-based runner.")
else:
    st.caption("No queued uploads right now.")

# -------------------------------
# Selection UI (always visible)
# -------------------------------
st.markdown("## Live Preview")
matched_codes = sorted(st.session_state.index.keys())
default_code = st.session_state.get("selected_code")

# If a code was previously selected, keep it, otherwise prefer first matched code, otherwise first from excel
if default_code in ad_codes:
    selected = default_code
elif matched_codes:
    selected = matched_codes[0]
else:
    selected = ad_codes[0] if ad_codes else ""

selected = st.selectbox(
    "Pick an Ad Code",
    options=ad_codes,
    index=ad_codes.index(selected) if selected in ad_codes else 0,
)

st.session_state.selected_code = selected

# Two-pane render
render_two_pane(df, ad_code_col, selected)

# Extra: quick list of matched codes for navigation
with st.expander("✅ Matched Ad Codes (so far)", expanded=False):
    if not matched_codes:
        st.caption("None yet.")
    else:
        st.write(", ".join(matched_codes[:200]) + ("…" if len(matched_codes) > 200 else ""))
