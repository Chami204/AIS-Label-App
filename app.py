import streamlit as st
import os
import zipfile
import tempfile
from utils.process_data import read_and_clean_excel
from utils.word_generator import generate_word_doc
from utils.barcode_generator import generate_barcodes

st.set_page_config(page_title="Pack Label Generator", page_icon="📦", layout="centered")
st.markdown("<h1 style='color:#023E8A;text-align:center;'>📦 Pack Label Generator</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    mfg = st.text_input("Enter Mfg. Date (e.g., 03-2025)")
    po_number = st.text_input("Enter PO Number")
    customer_name = st.text_input("Enter Customer Name")

    if mfg and po_number and customer_name:
        if st.button("Generate Word & Barcodes"):
            with tempfile.TemporaryDirectory() as tmpdir:
                df = read_and_clean_excel(uploaded_file)
                word_path = os.path.join(tmpdir, "Pack_Labels.docx")
                generate_word_doc(df, word_path, mfg, po_number, customer_name)

                barcode_dir = os.path.join(tmpdir, "barcodes")
                os.makedirs(barcode_dir, exist_ok=True)
                generate_barcodes(df, barcode_dir)

                # ZIP the barcodes
                zip_path = os.path.join(tmpdir, "barcodes.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(barcode_dir):
                        zipf.write(os.path.join(barcode_dir, file), file)

                st.success("✅ Generation completed!")
                st.download_button("Download Word Document", open(word_path, "rb"), file_name="Pack_Labels.docx")
                st.download_button("Download Barcode ZIP", open(zip_path, "rb"), file_name="barcodes.zip")
