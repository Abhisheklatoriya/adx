import streamlit as st
import docx
import re
import io

# 1. Page Configuration
st.set_page_config(page_title="Ad & Creative Matcher", layout="wide")

# 2. Custom CSS to improve the UI
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stExpander { background-color: white; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

def extract_data_from_docx(file):
    """Extracts all text and specifically looks for 8-digit Ad Codes."""
    doc = docx.Document(file)
    content = []
    for para in doc.paragraphs:
        if para.text.strip():
            content.append(para.text.strip())
    
    full_text = "\n".join(content)
    # Finds 8-digit numbers (common Ad Code format in your screenshots)
    ad_codes = sorted(list(set(re.findall(r'\b\d{8}\b', full_text))))
    return ad_codes, full_text

def main():
    st.title("🎯 Ad-Creative Matcher")
    st.info("Upload your Word document and your folder of assets. The app will automatically pair them based on the Ad Code.")

    # Sidebar for Uploads
    with st.sidebar:
        st.header("Step 1: Upload Files")
        word_file = st.file_uploader("Upload Ad Details (Word Doc)", type=['docx'])
        
        st.header("Step 2: Upload Assets")
        st.caption("Tip: You can select all files inside a folder (including subfolders)")
        creatives = st.file_uploader(
            "Upload Images/Videos", 
            accept_multiple_files=True,
            type=['png', 'jpg', 'jpeg', 'gif', 'mp4', 'mov', 'webm']
        )
        
        if creatives:
            st.success(f"Loaded {len(creatives)} assets.")

    # Main Processing Logic
    if word_file and creatives:
        ad_codes, full_doc_text = extract_data_from_docx(word_file)
        
        # Search & Filter bar
        search_query = st.text_input("🔍 Search by Ad Code", placeholder="e.g. 48725915")
        
        display_codes = [c for c in ad_codes if search_query in c] if search_query else ad_codes

        if not display_codes:
            st.warning("No matching Ad Codes found.")
            return

        for code in display_codes:
            # Find matching files in the uploaded list
            matched_files = [f for f in creatives if code in f.name]
            
            with st.expander(f"📦 Ad Code: {code} ({len(matched_files)} files found)", expanded=True):
                col1, col2 = st.columns([1, 1.5])
                
                with col1:
                    st.markdown("**Ad Details (From Doc):**")
                    # Try to find the specific block of text for this ad code
                    pattern = f"{code}.*?(?=\\n\\n|Ad Code:|$)"
                    match_text = re.search(pattern, full_doc_text, re.DOTALL)
                    if match_text:
                        st.code(match_text.group(0), language=None)
                    else:
                        st.write("Full details extraction unavailable. Showing Ad Code only.")

                with col2:
                    if matched_files:
                        for asset in matched_files:
                            file_ext = asset.name.split('.')[-1].lower()
                            
                            st.markdown(f"📄 `{asset.name}`")
                            
                            # Video Player
                            if file_ext in ['mp4', 'mov', 'webm']:
                                st.video(asset)
                            # Image Viewer
                            elif file_ext in ['jpg', 'jpeg', 'png', 'gif']:
                                st.image(asset, use_container_width=True)
                            
                            st.download_button(
                                label=f"Download {asset.name}",
                                data=asset,
                                file_name=asset.name,
                                mime=f"application/octet-stream",
                                key=f"dl_{asset.name}_{code}"
                            )
                            st.divider()
                    else:
                        st.error("No creative file found matching this Ad Code.")

    else:
        st.write("---")
        st.write("Please upload both the **Word Document** and the **Creative Assets** to begin.")

if __name__ == "__main__":
    main()
