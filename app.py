import pandas as pd
import zipfile
import base64
import os
import streamlit as st

def load_css(file_name):
    with open(file_name) as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

# Load custom CSS
load_css('style.css')

st.markdown('''
# **Excel Fusion - Combining sheets made simple! 🚀**
''', unsafe_allow_html=True)  # Add unsafe_allow_html=True to render HTML safely

# Excel file merge function
def excel_file_merge(zip_file_name):
    df_list = []
    archive = zipfile.ZipFile(zip_file_name, 'r')
    with zipfile.ZipFile(zip_file_name, "r") as f:
        for file in f.namelist():
            xlfile = archive.open(file)
            if file.endswith('.xlsx'):
                df_xl = pd.read_excel(xlfile, engine='openpyxl')
                df_xl['Note'] = file
                df_list.append(df_xl)
    df = pd.concat(df_list, ignore_index=True)
    return df

# Upload CSV data
with st.sidebar.header('1. Upload your ZIP file'):
    uploaded_file = st.sidebar.file_uploader("Please upload a zip file, containing all the excel sheets you wish to combine", type=["zip"])

# File download
def filedownload(df):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="merged_file.csv">Download Merged File as CSV</a>'
    return href

def xldownload(df):
    df.to_excel('data.xlsx', index=False)
    data = open('data.xlsx', 'rb').read()
    b64 = base64.b64encode(data).decode('UTF-8')
    href = f'<a href="data:file/xls;base64,{b64}" download="merged_file.xlsx">Download Merged File as XLSX</a>'
    return href

# Main panel
if st.sidebar.button('Submit'):
    if uploaded_file is not None:
        df = excel_file_merge(uploaded_file)
        st.header('**Voila! Here is your merged data!**')
        st.write(df)
        st.markdown(filedownload(df), unsafe_allow_html=True)
        st.markdown(xldownload(df), unsafe_allow_html=True)
    else:
        st.error('Please upload a ZIP file first.')
else:
    st.info('Awaiting for ZIP file to be uploaded.')

# Additional text above footer
st.markdown('''
---

#### Struggling to manage multiple Excel files? Meet ExcelFusion, the simple tool that brings order to your data. With just a few clicks, you can combine all your separate Excel sheets into one organized file!
''', unsafe_allow_html=True)

# Footer
footer_html = """
<div class="footer">
<p>Created with ❤️ by Anusha Tomar | <a href="https://github.com" target="_blank">GitHub Link</a> | © 2024 ExcelFusion</p>
</div>
"""
st.markdown(footer_html, unsafe_allow_html=True)