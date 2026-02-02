import streamlit as st
import docx
import re
import zipfile
import io

# 1. Expand limits and set layout
st.set_page_config(page_title="Ad Matcher High-Speed", layout="wide")

# 2. Cache the Word Document extraction
@st.cache_data
def get_ad_data(file):
    doc = docx.Document(file)
    text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    codes = sorted(list(set(re.findall(r'\b\d{8}\b', text))))
    return codes, text

# 3. Cache the Assets in memory so they don't reload on every click
@st.cache_resource
def load_assets(uploaded_files):
    processed_assets = []
    for uploaded_file in uploaded_files:
        if uploaded_file.name.lower().endswith('.zip'):
            with zipfile.ZipFile(uploaded_file) as z:
                for file_info in z.infolist():
                    if file_info.is_dir() or "__MACOSX" in file_info.filename:
                        continue
                    # Read into memory
                    with z.open(file_info) as f:
                        data = f.read()
                        processed_assets.append({
                            "name": file_info.filename,
                            "data": data,
                            "ext": file_info.filename.split('.')[-1].lower()
                        })
        else:
            # Handle individual files
            processed_assets.append({
                "name": uploaded_file.name,
                "data": uploaded_file.getvalue(),
                "ext": uploaded_file.name.split('.')[-1].lower()
            })
    return processed_assets

st.title("⚡ Ultra-Fast Ad Matcher")

# Sidebar Uploads
with st.sidebar:
    st.header("Upload")
    word_file = st.file_uploader("Upload Word Doc", type=['docx'])
    raw_files = st.file_uploader("Upload Assets or ZIP", accept_multiple_files=True)
    if st.button("Clear Cache / Reset"):
        st.cache_resource.clear()
        st.rerun()

if word_file and raw_files:
    # This runs once and stays in memory
    all_assets = load_assets(raw_files)
    ad_codes, full_text = get_ad_data(word_file)

    st.success(f"Loaded {len(all_assets)} files. Found {len(ad_codes)} Ad Codes.")
    
    # Fast Search
    search = st.text_input("🔍 Quick Search Ad Code", placeholder="Type to filter...")
    filtered_codes = [c for c in ad_codes if search in c] if search else ad_codes

    for code in filtered_codes:
        # Match based on filename containing the code
        matches = [a for a in all_assets if code in a['name']]
        
        if matches:
            with st.expander(f"✅ Ad {code} - {len(matches)} files found", expanded=True):
                c1, c2 = st.columns([1, 1.5])
                
                with c1:
                    st.markdown("**Ad Specs:**")
                    pattern = f"{code}.*?(?=\\n\\n|Ad Code:|$)"
                    found = re.search(pattern, full_text, re.DOTALL)
                    st.code(found.group(0) if found else f"Ad Code: {code}", language=None)
                
                with c2:
                    for asset in matches:
                        st.caption(f"File: {asset['name']}")
                        # Media logic
                        if asset['ext'] in ['mp4', 'mov', 'webm']:
                            st.video(asset['data'])
                        elif asset['ext'] in ['mp3', 'wav']:
                            st.audio(asset['data'])
                        elif asset['ext'] in ['jpg', 'jpeg', 'png', 'gif']:
                            st.image(asset['data'])
                        
                        st.download_button("Download", data=asset['data'], file_name=asset['name'], key=f"dl_{asset['name']}_{code}")
                        st.divider()

else:
    st.info("Please upload your Word doc and creative assets to begin.")
