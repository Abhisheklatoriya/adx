import streamlit as st
import docx
import re
import os

def extract_ad_details(uploaded_word_file):
    doc = docx.Document(uploaded_word_file)
    full_text = [para.text for para in doc.paragraphs]
    # Logic to group text by "Ad Code"
    return "\n".join(full_text)

st.title("🎯 Ad & Creative Matcher")

# 1. File Uploaders
word_file = st.file_uploader("Upload Ad Details (Word Doc)", type=['docx'])
creatives = st.file_uploader("Upload Creative Assets", type=['png', 'jpg', 'mp4', 'mov'], accept_multiple_files=True)

if word_file and creatives:
    # Example logic: searching for the ID 48725915 in filenames
    ad_id = "48725915" # This would be parsed from your Doc
    
    for asset in creatives:
        if ad_id in asset.name:
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Ad Details")
                st.text(f"Ad Code: {ad_id}\nBrand: Bell WiFi\nOutlet: RedFlagDeals")
                
            with col2:
                st.subheader("Creative Preview")
                if asset.name.endswith(('.mp4', '.mov')):
                    st.video(asset)
                else:
                    st.image(asset)
                
                st.download_button("Download Matched Asset", data=asset, file_name=asset.name)
