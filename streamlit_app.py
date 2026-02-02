import streamlit as st
import docx
import re

# 1. Page Configuration
st.set_page_config(page_title="Ad & Creative Matcher", layout="wide")

def extract_data_from_docx(file):
    doc = docx.Document(file)
    content = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    full_text = "\n".join(content)
    # Finds 8-digit numbers (Ad Codes)
    ad_codes = sorted(list(set(re.findall(r'\b\d{8}\b', full_text))))
    return ad_codes, full_text

def main():
    st.title("🎯 Ad-Creative Matcher")

    with st.sidebar:
        st.header("Step 1: Upload Files")
        word_file = st.file_uploader("Upload Ad Details (Word Doc)", type=['docx'])
        
        st.header("Step 2: Upload Assets")
        # Added 'mp3' and 'wav' to the type list
        creatives = st.file_uploader(
            "Upload Images/Videos/Audio", 
            accept_multiple_files=True,
            type=['png', 'jpg', 'jpeg', 'gif', 'mp4', 'mov', 'webm', 'mp3', 'wav']
        )
        
    if word_file and creatives:
        ad_codes, full_doc_text = extract_data_from_docx(word_file)
        
        search_query = st.text_input("🔍 Search by Ad Code", placeholder="e.g. 48725915")
        display_codes = [c for c in ad_codes if search_query in c] if search_query else ad_codes

        for code in display_codes:
            matched_files = [f for f in creatives if code in f.name]
            
            with st.expander(f"📦 Ad Code: {code} ({len(matched_files)} files found)", expanded=True):
                col1, col2 = st.columns([1, 1.5])
                
                with col1:
                    st.markdown("**Ad Details:**")
                    pattern = f"{code}.*?(?=\\n\\n|Ad Code:|$)"
                    match_text = re.search(pattern, full_doc_text, re.DOTALL)
                    if match_text:
                        st.code(match_text.group(0), language=None)
                
                with col2:
                    if matched_files:
                        for asset in matched_files:
                            ext = asset.name.split('.')[-1].lower()
                            st.markdown(f"📄 `{asset.name}`")
                            
                            # --- AUDIO SUPPORT ---
                            if ext in ['mp3', 'wav']:
                                st.audio(asset)
                            
                            # --- VIDEO SUPPORT ---
                            elif ext in ['mp4', 'mov', 'webm']:
                                st.video(asset)
                            
                            # --- IMAGE SUPPORT ---
                            elif ext in ['jpg', 'jpeg', 'png', 'gif']:
                                st.image(asset, use_container_width=True)
                            
                            st.download_button(
                                label=f"Download {asset.name}",
                                data=asset,
                                file_name=asset.name,
                                key=f"dl_{asset.name}_{code}"
                            )
                            st.divider()
                    else:
                        st.error("No creative file found matching this Ad Code.")
    else:
        st.info("Waiting for Word Doc and Creative Assets...")

if __name__ == "__main__":
    main()
