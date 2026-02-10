import re
import zipfile
import shutil
import tempfile
from pathlib import Path
from typing import Optional, Set, List, Tuple

import streamlit as st
import pandas as pd

st.set_page_config(page_title="Ad Creative Matcher", layout="wide")
st.title("⚡ Ad Creative Matcher (Excel + Live Processing)")

CODE_RE = re.compile(r"\b(\d{8})\b")


def ensure_workspace() -> Path:
    if "workspace_dir" not in st.session_state:
        st.session_state.workspace_dir = tempfile.mkdtemp(prefix="admatcher_")
    return Path(st.session_state.workspace_dir)


def output_dir() -> Path:
    out = ensure_workspace() / "output"
    out.mkdir(parents=True, exist_ok=True)
    return out


def reset_workspace() -> None:
    if "workspace_dir" in st.session_state:
        shutil.rmtree(st.session_state.workspace_dir, ignore_errors=True)
        del st.session_state["workspace_dir"]
    if "processed_keys" in st.session_state:
        del st.session_state["processed_keys"]
    st.cache_data.clear()


@st.cache_data
def load_excel_auto_header(excel_file) -> Tuple[pd.DataFrame, str, List[str], Set[str]]:
    # Read raw so we can locate the real header row (your sheet has title rows above)
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

    # Re-read using detected header row
    df = pd.read_excel(excel_file, header=header_row, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    ad_code_col = None
    for c in df.columns:
        cl = c.lower().strip()
        if "ad" in cl and "code" in cl:
            ad_code_col = c
            break

    if not ad_code_col:
        raise ValueError(f"Header found, but no Ad Code column. Columns: {list(df.columns)}")

    df[ad_code_col] = df[ad_code_col].astype(str).str.strip()
    df = df[df[ad_code_col].str.match(r"^\d{8}$")]

    ad_codes = df[ad_code_col].tolist()
    ad_codes_set = set(ad_codes)
    return df, ad_code_col, ad_codes, ad_codes_set


def safe_relpath(p: str) -> str:
    p = p.replace("\\", "/")
    p = p.lstrip("/").lstrip("./")
    return p


def extract_ad_code_from_filename(name: str, valid_codes: Set[str]) -> Optional[str]:
    for m in CODE_RE.findall(name):
        if m in valid_codes:
            return m
    return None


def copy_stream_to_disk(src, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as out:
        shutil.copyfileobj(src, out, length=1024 * 1024)  # 1MB chunks


def list_matches_for_code(code: str) -> List[Path]:
    base = output_dir() / code
    if not base.exists():
        return []
    return sorted([p for p in base.rglob("*") if p.is_file()])


def file_ext(p: Path) -> str:
    return p.suffix.lower().lstrip(".")


def read_preview_bytes(p: Path, max_mb: int = 25) -> Optional[bytes]:
    if p.stat().st_size > max_mb * 1024 * 1024:
        return None
    return p.read_bytes()


def process_new_uploads(uploaded_files, valid_codes: Set[str]) -> None:
    if "processed_keys" not in st.session_state:
        st.session_state.processed_keys = set()

    # Estimate steps
    steps = 0
    items = []
    for uf in uploaded_files:
        key = f"{uf.name}|{getattr(uf, 'size', 'na')}"
        items.append((uf, key))
        if key in st.session_state.processed_keys:
            continue

        if uf.name.lower().endswith(".zip"):
            try:
                uf.seek(0)
                with zipfile.ZipFile(uf) as z:
                    entries = [
                        i for i in z.infolist()
                        if (not i.is_dir()) and "__MACOSX" not in i.filename
                    ]
                steps += max(len(entries), 1)
            except Exception:
                steps += 1
        else:
            steps += 1

    if steps == 0:
        st.info("No new uploads to process.")
        return

    progress = st.progress(0.0)
    status = st.empty()

    done = 0
    out_root = output_dir()

    for uf, key in items:
        if key in st.session_state.processed_keys:
            continue

        status.write(f"Processing **{uf.name}** …")

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
                            inner_name = safe_relpath(info.filename)
                            code = extract_ad_code_from_filename(inner_name, valid_codes)

                            if code:
                                dest = out_root / code / inner_name
                                with z.open(info) as src:
                                    copy_stream_to_disk(src, dest)

                            done += 1
                            progress.progress(min(done / steps, 1.0))

            except Exception as e:
                st.warning(f"Could not read ZIP {uf.name}: {e}")
                done += 1
                progress.progress(min(done / steps, 1.0))

        else:
            fname = safe_relpath(uf.name)
            code = extract_ad_code_from_filename(fname, valid_codes)
            if code:
                dest = out_root / code / fname
                try:
                    uf.seek(0)
                    copy_stream_to_disk(uf, dest)
                except Exception as e:
                    st.warning(f"Could not save {uf.name}: {e}")

            done += 1
            progress.progress(min(done / steps, 1.0))

        st.session_state.processed_keys.add(key)

    status.write("✅ Finished processing currently uploaded files.")


with st.sidebar:
    st.header("Upload")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    assets_files = st.file_uploader("Upload Assets or ZIPs", accept_multiple_files=True)

    st.divider()
    if st.button("Reset / Clear Session"):
        reset_workspace()
        st.rerun()

if not excel_file:
    st.info("Upload the Excel ad list to begin.")
    st.stop()

try:
    df, ad_code_col, ad_codes, ad_codes_set = load_excel_auto_header(excel_file)
except Exception as e:
    st.error(f"Excel load error: {e}")
    st.stop()

st.success(f"Loaded **{len(df)}** ads. Detected Ad Code column: **{ad_code_col}**")

if assets_files:
    process_new_uploads(assets_files, ad_codes_set)
else:
    st.info("Upload creatives/ZIPs — matches will appear as soon as files are processed.")

search = st.text_input("🔍 Search Ad Code", placeholder="Type partial or full 8-digit ad code…")
filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes
only_matched = st.checkbox("Show only ad codes with matches", value=True)

for code in filtered_codes:
    matches = list_matches_for_code(code)
    if only_matched and not matches:
        continue

    with st.expander(f"{'✅' if matches else '⚪️'} Ad {code} — {len(matches)} files", expanded=bool(matches)):
        c1, c2 = st.columns([1, 1.6])

        with c1:
            st.markdown("**Ad Specs (Excel row)**")
            st.dataframe(df[df[ad_code_col] == code], use_container_width=True)

        with c2:
            if not matches:
                st.caption("No matching files yet.")
            else:
                for p in matches:
                    rel = p.relative_to(output_dir() / code)
                    st.caption(str(rel))

                    ext = file_ext(p)
                    preview = read_preview_bytes(p, max_mb=25)

                    if preview is None:
                        st.info("Preview skipped (file is large). Download to view.")
                    else:
                        if ext in ["mp4", "mov", "webm"]:
                            st.video(preview)
                        elif ext in ["mp3", "wav"]:
                            st.audio(preview)
                        elif ext in ["jpg", "jpeg", "png", "gif"]:
                            st.image(preview)

                    st.download_button(
                        "Download",
                        data=p.read_bytes(),
                        file_name=p.name,
                        key=f"dl_{code}_{p.as_posix()}",
                    )
                    st.divider()
