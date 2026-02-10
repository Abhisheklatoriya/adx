import streamlit as st
import pandas as pd
import zipfile
import re
import shutil
import tempfile
from pathlib import Path
from typing import Optional, Set, List

# -------------------------------
# App Config
# -------------------------------
st.set_page_config(page_title="Ad Creative Matcher", layout="wide")
st.title("⚡ Ad Matcher (Excel + Live Processing + Disk Storage)")

# -------------------------------
# Excel loader
# -------------------------------
@st.cache_data
import pandas as pd
import streamlit as st

@st.cache_data
def load_excel(file):
    # 1) Read raw (no header) so we can find the real header row
    raw = pd.read_excel(file, header=None, engine="openpyxl")

    # 2) Find the row that contains "Ad Code" (case-insensitive)
    header_row = None
    for i in range(min(50, len(raw))):
        row_vals = raw.iloc[i].astype(str).str.strip().str.lower()
        if row_vals.eq("ad code").any() or row_vals.str.contains(r"\bad\s*code\b", regex=True).any():
            header_row = i
            break

    if header_row is None:
        # Helpful debug: show what the first rows look like
        raise ValueError(
            "Could not find the table header row containing 'Ad Code'. "
            "Try increasing the scan range or confirm the sheet format."
        )

    # 3) Re-read using the detected header row
    df = pd.read_excel(file, header=header_row, engine="openpyxl")

    # 4) Normalize column names
    df.columns = [str(c).strip() for c in df.columns]

    # 5) Find Ad Code column robustly (handles 'Ad Code', 'AdCode', 'Ad  Code', etc.)
    lower_cols = {c.lower(): c for c in df.columns}
    ad_code_col = None
    for c in df.columns:
        cl = c.lower().strip()
        if "ad" in cl and "code" in cl:
            ad_code_col = c
            break

    if not ad_code_col:
        raise ValueError(f"Found header row, but couldn't find Ad Code column. Columns seen: {list(df.columns)}")

    # 6) Clean & keep only 8-digit codes
    df[ad_code_col] = df[ad_code_col].astype(str).str.strip()
    df = df[df[ad_code_col].str.match(r"^\d{8}$")]

    return df, ad_code_col


# -------------------------------
# Disk workspace (per session)
# -------------------------------
def ensure_workspace() -> Path:
    if "workspace_dir" not in st.session_state:
        st.session_state.workspace_dir = tempfile.mkdtemp(prefix="admatcher_")
    return Path(st.session_state.workspace_dir)

def get_output_dir() -> Path:
    base = ensure_workspace()
    out = base / "output"
    out.mkdir(parents=True, exist_ok=True)
    return out

# -------------------------------
# Helpers
# -------------------------------
CODE_RE = re.compile(r"\b(\d{8})\b")

def extract_code_from_name(name: str, valid_codes_set: Set[str]) -> Optional[str]:
    for m in CODE_RE.findall(name):
        if m in valid_codes_set:
            return m
    return None

def safe_filename(p: str) -> str:
    p = p.replace("\\", "/")
    p = p.lstrip("/").lstrip("./")
    return p

def copy_stream_to_disk(src_fileobj, dest_path: Path):
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    with open(dest_path, "wb") as out:
        shutil.copyfileobj(src_fileobj, out, length=1024 * 1024)  # 1MB chunks

# -------------------------------
# Incremental processing
# -------------------------------
def process_new_uploads(uploaded_files, valid_codes_set: Set[str]) -> None:
    if "processed_keys" not in st.session_state:
        st.session_state.processed_keys = set()

    out_dir = get_output_dir()

    progress = st.progress(0.0)
    status = st.empty()

    # Rough steps: files + zip entries
    steps = 0
    for uf in uploaded_files:
        key = f"{uf.name}|{getattr(uf, 'size', 'na')}"
        if key in st.session_state.processed_keys:
            continue

        if uf.name.lower().endswith(".zip"):
            try:
                uf.seek(0)
                with zipfile.ZipFile(uf) as z:
                    entries = [i for i in z.infolist() if (not i.is_dir()) and "__MACOSX" not in i.filename]
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
        if key in st.session_state.processed_keys:
            continue

        status.write(f"Processing **{uf.name}** …")

        if uf.name.lower().endswith(".zip"):
            try:
                uf.seek(0)
                with zipfile.ZipFile(uf) as z:
                    entries = [i for i in z.infolist() if (not i.is_dir()) and "__MACOSX" not in i.filename]
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
                                    copy_stream_to_disk(src, dest)

                            done += 1
                            progress.progress(min(done / steps, 1.0))
            except Exception as e:
                st.warning(f"Could not read ZIP {uf.name}: {e}")
                done += 1
                progress.progress(min(done / steps, 1.0))

        else:
            fname = safe_filename(uf.name)
            code = extract_code_from_name(fname, valid_codes_set)
            if code:
                dest = out_dir / code / fname
                try:
                    uf.seek(0)
                    copy_stream_to_disk(uf, dest)
                except Exception as e:
                    st.warning(f"Could not save file {uf.name}: {e}")

            done += 1
            progress.progress(min(done / steps, 1.0))

        st.session_state.processed_keys.add(key)

    status.write("✅ Done processing currently uploaded files.")

def list_matches_for_code(code: str) -> List[Path]:
    out_dir = get_output_dir() / code
    if not out_dir.exists():
        return []
    return sorted([p for p in out_dir.rglob("*") if p.is_file()])

def ext_of(path: Path) -> str:
    return path.suffix.lower().lstrip(".")

def read_small_preview(path: Path, max_mb: int = 25) -> Optional[bytes]:
    max_bytes = max_mb * 1024 * 1024
    if path.stat().st_size > max_bytes:
        return None
    return path.read_bytes()

# -------------------------------
# Sidebar
# -------------------------------
with st.sidebar:
    st.header("Upload")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    raw_files = st.file_uploader("Upload Assets or ZIPs", accept_multiple_files=True)

    if st.button("Reset session"):
        if "workspace_dir" in st.session_state:
            shutil.rmtree(st.session_state.workspace_dir, ignore_errors=True)
        for k in ["workspace_dir", "processed_keys"]:
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

                    preview_bytes = read_small_preview(p, max_mb=25)
                    ext = ext_of(p)

                    if preview_bytes is None:
                        st.info("Preview skipped (file is large). Download to view.")
                    else:
                        if ext in ["mp4", "mov", "webm"]:
                            st.video(preview_bytes)
                        elif ext in ["mp3", "wav"]:
                            st.audio(preview_bytes)
                        elif ext in ["jpg", "jpeg", "png", "gif"]:
                            st.image(preview_bytes)

                    st.download_button(
                        "Download",
                        data=p.read_bytes(),  # note: for very large files, we can optimize this next
                        file_name=p.name,
                        key=f"dl_{code}_{p.as_posix()}"
                    )
                    st.divider()
