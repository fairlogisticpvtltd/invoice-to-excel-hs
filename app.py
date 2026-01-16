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
    Improved parser:
    - Ignores header / title lines
    - Detects Description/Item, Qty/Quantity, Origin
    - Extracts only line items
    """
    rows = []
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    header_keywords = [
        "invoice", "tax", "bill to", "ship to", "date",
        "total", "subtotal", "amount in words", "bank"
    ]

    for line in lines:
        low = line.lower()

        # Skip header/footer info
        if any(h in low for h in header_keywords):
            continue

        # Likely item lines (contain qty or numbers + text)
        if any(k in low for k in ["qty", "quantity"]) or any(c.isdigit() for c in line):
            rows.append({
                "Full Description": line,
                "Brand": "",
                "Model": "",
                "Size": "",
                "Qty": "",
                "Packing": "",
                "Origin": "",
                "Total Price": ""
            })

    return pd.DataFrame(rows)


def map_hs_codes(invoice_df, hs_df):
    # Normalize column names
    hs_df.columns = [c.strip().lower() for c in hs_df.columns]

    # Auto-detect HS code and description columns
    desc_col = None
    hs_col = None

    for col in hs_df.columns:
        if "desc" in col:
            desc_col = col
        if "hs" in col:
            hs_col = col

    if not desc_col or not hs_col:
        raise ValueError("HS code file must contain description and hs code columns")

    # Detect optional unit column
    unit_col = None
    for col in hs_df.columns:
        if "unit" in col:
            unit_col = col

    select_cols = [desc_col, hs_col]
    if unit_col:
        select_cols.append(unit_col)

    hs_map = hs_df[select_cols]

    hs_codes = []
    hs_descs = []
    hs_units = []

    for desc in invoice_df["Full Description"]:
        # Pre-clean invoice description for better HS matching
        clean_desc = desc.lower()
        keywords = [
            "elbow", "tee", "coupling", "adapter", "union",
            "pipe", "pvc", "fitting", "valve"
        ]
        for k in keywords:
            if k in clean_desc:
                clean_desc += f" {k}"

        match = process.extractOne(
            clean_desc,
            hs_map[desc_col].str.lower(),
            scorer=fuzz.token_set_ratio
        )
        if match and match[1] > 70:
            row = hs_map.iloc[match[2]]
            hs_codes.append(row[hs_col])
            hs_descs.append(row[desc_col])
            hs_units.append(row[unit_col] if unit_col else "")
        else:
            hs_codes.append("NOT FOUND")
            hs_descs.append("NOT FOUND")
            hs_units.append("")

    invoice_df["HS Code"] = hs_codes
    invoice_df["HS Description"] = hs_descs
    invoice_df["Unit"] = hs_units
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
