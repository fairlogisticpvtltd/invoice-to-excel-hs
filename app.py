# Invoice to Excel with HS Code Mapping Program
# ------------------------------------------
# This is a complete Python program that:
# 1. Uploads an invoice (PDF or image)
# 2. Extracts invoice line-item data using OCR
# 3. Structures data into columns
# 4. Maps HS Codes using a provided HS code Excel file
# 5. Exports the final result to Excel

# REQUIREMENTS:
# pip install pytesseract pdfplumber pandas openpyxl pillow rapidfuzz streamlit
# Install Tesseract OCR engine separately (system-level)

import streamlit as st
import pdfplumber
import pandas as pd
from PIL import Image
import pytesseract
from rapidfuzz import process, fuzz

# ---------------------- CONFIG ----------------------
st.set_page_config(page_title="Invoice to Excel Converter", layout="wide")
st.title("📄 Invoice to Excel Converter with HS Code Mapping")

# ---------------------- FUNCTIONS ----------------------
def extract_text_from_pdf(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    return text


def extract_text_from_image(image_file):
    image = Image.open(image_file)
    return pytesseract.image_to_string(image)


def parse_invoice_text(text):
    """
    VERY IMPORTANT:
    This is a semi-automatic parser.
    You may need to fine-tune rules based on your invoice format.
    """
    rows = []
    lines = text.split("\n")

    for line in lines:
        if any(char.isdigit() for char in line):
            rows.append({
                "Full Description": line,
                "Brand": "",
                "Model": "",
                "Size": "",
                "Qty": "",
                "Packing": "",
                "Total Price": ""
            })
    return pd.DataFrame(rows)


def map_hs_codes(invoice_df, hs_df):
    hs_df.columns = [c.lower() for c in hs_df.columns]
    hs_map = hs_df[["description", "hs_code"]]

    hs_codes = []
    for desc in invoice_df["Full Description"]:
        match, score, idx = process.extractOne(
            desc, hs_map["description"], scorer=fuzz.token_sort_ratio
        )
        if score > 70:
            hs_codes.append(hs_map.iloc[idx]["hs_code"])
        else:
            hs_codes.append("NOT FOUND")

    invoice_df["HS Code"] = hs_codes
    return invoice_df

# ---------------------- UI ----------------------
invoice_file = st.file_uploader("Upload Invoice (PDF or Image)", type=["pdf", "png", "jpg", "jpeg"])
hs_file = st.file_uploader("Upload HS Code Excel File", type=["xlsx"])

if invoice_file and hs_file:
    if invoice_file.type == "application/pdf":
        raw_text = extract_text_from_pdf(invoice_file)
    else:
        raw_text = extract_text_from_image(invoice_file)

    st.subheader("📄 Extracted Text Preview")
    st.text(raw_text[:3000])

    invoice_df = parse_invoice_text(raw_text)

    hs_df = pd.read_excel(hs_file)
    final_df = map_hs_codes(invoice_df, hs_df)

    st.subheader("📊 Invoice Table")
    st.dataframe(final_df, use_container_width=True)

    output_file = "invoice_with_hs_codes.xlsx"
    final_df.to_excel(output_file, index=False)

    with open(output_file, "rb") as f:
        st.download_button("⬇ Download Excel File", f, file_name=output_file)

# ---------------------- NOTES ----------------------
# - HS Excel file MUST have columns:
#   description | hs_code
# - Parsing rules can be customized per supplier
# - Accuracy improves with clean invoices
# - This can be upgraded with AI/NLP later
