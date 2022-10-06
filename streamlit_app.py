import preprocessing
import streamlit as st

st.header('D. Query tool processor (accept XLSX file only)')
uploaded_file = st.file_uploader('Upload here', type=['xlsx'])
if uploaded_file is not None:

    # to advise user to wait while preprocessing is ongoing
    st.header('Please wait patiently until the Download Button appears')

    # preprocess the file
    preprocessing.main(uploaded_file)

    # download button to download the processed zip file
    with open("output.zip", "rb") as file_pointer:
        btn = st.download_button(
            label="Download Processed ZIP file",
            data=file_pointer,
            file_name="output.zip",
            mime="application/zip"
        )
