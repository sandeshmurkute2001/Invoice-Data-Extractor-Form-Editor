import os
import re
import json
import hashlib
import openai
import pandas as pd
import pdfplumber
import streamlit as st
import time

# ========== CONFIG ==========
OPENAI_KEY = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"  # Replace with your key
client = openai.OpenAI(api_key=OPENAI_KEY)

output_excel_path = "extracted_invoice_data.xlsx"
uploaded_pdf_dir = "uploaded_pdfs"
os.makedirs(uploaded_pdf_dir, exist_ok=True)

# ========== UTILITY FUNCTIONS ==========

def hash_text(text):
    return hashlib.sha256(text.encode('utf-8')).hexdigest()

def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:2]:  # First 2 pages max
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    text = re.sub(r'(?<=\d)(Nos|KGS|PCS)', r'\n\1', text)
    text = re.sub(r'(Nos|KGS|PCS)(?=\d)', r'\1\n', text)
    text = re.sub(r'(?<=\d)(?=\d{10})', '\n', text)
    return text

def extract_data_using_gpt(ocr_text, invoice_file_name):
    prompt = f'''
You are an AI assistant helping extract structured data from OCR text of an Indian tax invoice.

Your task is to extract all invoice fields in a JSON format. Include the following fields:
- InvoiceNumber
- InvoiceDate
- PO_Number
- BuyerDate
- SellerName
- SellerAddress
- SellerGST
- BuyerName
- BuyerAddress
- BuyerGST
- Items (a list of line items with fields: HSN/SAC, ItemNumber, ItemDescription, ItemQuantity, ItemRate, ItemAmount, GST%, GSTAmount)
- GSTAmount
- TotalAmountAfterTax

OCR TEXT:
{ocr_text}

Only return valid pure JSON. No explanations.
'''
    response = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content": "You are a reliable invoice data extraction engine."},
            {"role": "user", "content": prompt}
        ]
    )

    response_text = response.choices[0].message.content.strip()
    json_match = re.search(r"\{.*\}", response_text, re.DOTALL)
    if not json_match:
        raise ValueError("Could not extract JSON from GPT response.")
    return json.loads(json_match.group(0))

# ========== STREAMLIT UI ==========

st.title("Invoice Data Extractor")

uploaded_file = st.file_uploader("Upload a PDF Invoice", type=["pdf"])

if uploaded_file:
    file_path = os.path.join(uploaded_pdf_dir, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success(f"Uploaded {uploaded_file.name}")

    with st.spinner("Extracting text and analyzing with GPT..."):
        try:
            ocr_text = extract_text_from_pdf(file_path)
            data = extract_data_using_gpt(ocr_text, uploaded_file.name)
            time.sleep(1.5)

            # Show structured info in UI
            st.subheader("Invoice Metadata")
            st.json({
                "Invoice Number": data.get("InvoiceNumber", ""),
                "Invoice Date": data.get("InvoiceDate", ""),
                "PO Number": data.get("PO_Number", ""),
                "Buyer Date": data.get("BuyerDate", "")
            })

            st.subheader("Seller Info")
            st.json({
                "Name": data.get("SellerName", ""),
                "Address": data.get("SellerAddress", ""),
                "GST": data.get("SellerGST", "")
            })

            st.subheader("Buyer Info")
            st.json({
                "Name": data.get("BuyerName", ""),
                "Address": data.get("BuyerAddress", ""),
                "GST": data.get("BuyerGST", "")
            })

            st.subheader("Line Items")
            df_items = pd.DataFrame(data["Items"])
            st.dataframe(df_items)

            st.subheader("Summary")
            st.json({
                "GST Amount": data.get("GSTAmount", ""),
                "Total Amount After Tax": data.get("TotalAmountAfterTax", "")
            })

            # Export to Excel
            output_data = []
            for item in data["Items"]:
                output_data.append({
                    "Invoice Name": uploaded_file.name,
                    "InvoiceNumber": data.get("InvoiceNumber", ""),
                    "InvoiceDate": data.get("InvoiceDate", ""),
                    "PO_Number": data.get("PO_Number", ""),
                    "Buyer Date": data.get("BuyerDate", ""),
                    "SellerName": data.get("SellerName", ""),
                    "SellerAddress": data.get("SellerAddress", ""),
                    "SellerGST": data.get("SellerGST", ""),
                    "BuyerName": data.get("BuyerName", ""),
                    "BuyerAddress": data.get("BuyerAddress", ""),
                    "BuyerGST": data.get("BuyerGST", ""),
                    "HSN/SAC": item.get("HSN/SAC", ""),
                    "Item Number": item.get("ItemNumber", ""),
                    "Item Description": item.get("ItemDescription", ""),
                    "Item Quantity": item.get("ItemQuantity", ""),
                    "Item Rate": item.get("ItemRate", ""),
                    "Item Amount": item.get("ItemAmount", ""),
                    "GST%": item.get("GST%", ""),
                    "GST Amount": data.get("GSTAmount", ""),
                    "Total Amount After Tax": data.get("TotalAmountAfterTax", "")
                })
            df = pd.DataFrame(output_data)
            df.to_excel(output_excel_path, index=False)
            st.success(f"Excel saved to: {output_excel_path}")

        except Exception as e:
            st.error(f"Error: {e}")
