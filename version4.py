import os
import re
import json
import base64
import time
import openai
import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import load_workbook

# ========== CONFIG ==========
OPENAI_KEY = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
client = openai.OpenAI(api_key=OPENAI_KEY)

output_folder = r"C:\Python_Practice\UI_Extraction\UI_Output"
os.makedirs(output_folder, exist_ok=True)
output_excel_path = os.path.join(output_folder, "corrected_invoice_data.xlsx")

uploaded_pdf_dir = "uploaded_pdfs"
os.makedirs(uploaded_pdf_dir, exist_ok=True)

# ========== UTILITY FUNCTIONS ==========
def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:2]:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

def extract_data_using_gpt(ocr_text, invoice_file_name):
    prompt = f'''
You are an AI assistant helping extract structured data from OCR text of an Indian tax invoice.

Please extract the following invoice-level and line-item fields from the OCR text. Follow these strict instructions:

1. Parse line items carefully.
2. If the quantity and rate appear merged like "110.0012 Nos", interpret them as:
   - ItemRate: 110.00
   - ItemQuantity: 12
   (Rate comes before the quantity and unit)
3. Always compute:
   - ItemAmount = ItemRate × ItemQuantity (and extract if available)
   - Extract accurate GST% from line if available
   - If GST% is not visible, calculate it using: GST% = (GSTAmount / (TotalAmountAfterTax - GSTAmount)) * 100
4. Do not guess or hallucinate; only use visible or derivable data.
5. If a value is not visible and cannot be computed, leave it as "".

Extract the following in valid JSON:

{{
  "InvoiceName": "{invoice_file_name}",
  "InvoiceNumber": "...",
  "InvoiceDate": "...",
  "PO_Number": "...",
  "BuyerDate": "...",
  "SellerName": "...",
  "SellerAddress": "...",
  "SellerGST": "...",
  "BuyerName": "...",
  "BuyerAddress": "...",
  "BuyerGST": "...",
  "GSTAmount": "...",
  "Items": [
    {{
      "HSN/SAC": "...",
      "ItemNumber": "...",
      "ItemDescription": "...",
      "ItemQuantity": "...",
      "ItemRate": "...",
      "ItemAmount": "...",
      "GST%": "..."
    }}
  ],
  "TotalAmountAfterTax": "..."
}}

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

def show_pdf_scrollable(pdf_path):
    with open(pdf_path, "rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode("utf-8")
    pdf_display = f'''
        <iframe src="data:application/pdf;base64,{base64_pdf}" 
                width="100%" height="1000px" 
                style="border: 1px solid #ccc; border-radius: 10px;">
        </iframe>
    '''
    st.markdown(pdf_display, unsafe_allow_html=True)

# ========== STREAMLIT UI ==========
st.set_page_config(layout="wide")
st.title("Invoice Data Extractor & Form Editor")

uploaded_file = st.file_uploader("Upload a PDF Invoice", type=["pdf"])

if uploaded_file:
    file_path = os.path.join(uploaded_pdf_dir, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success(f"Uploaded {uploaded_file.name}")

    with st.spinner("Extracting and analyzing invoice..."):
        try:
            ocr_text = extract_text_from_pdf(file_path)
            data = extract_data_using_gpt(ocr_text, uploaded_file.name)
            time.sleep(1)

            st.header("Extracted Invoice Data & Scrollable Invoice Preview")
            col1, col2 = st.columns([2, 3])

            with col1:
                st.subheader("🔧 Review and Edit Invoice Data")

                invoice_number = st.text_input("Invoice Number", data.get("InvoiceNumber", ""))
                invoice_date = st.text_input("Invoice Date", data.get("InvoiceDate", ""))
                po_number = st.text_input("PO Number", data.get("PO_Number", ""))
                buyer_date = st.text_input("Buyer Date", data.get("BuyerDate", ""))

                st.subheader("Seller Information")
                seller_name = st.text_input("Seller Name", data.get("SellerName", ""))
                seller_address = st.text_area("Seller Address", data.get("SellerAddress", ""))
                seller_gst = st.text_input("Seller GST", data.get("SellerGST", ""))

                st.subheader("Buyer Information")
                buyer_name = st.text_input("Buyer Name", data.get("BuyerName", ""))
                buyer_address = st.text_area("Buyer Address", data.get("BuyerAddress", ""))
                buyer_gst = st.text_input("Buyer GST", data.get("BuyerGST", ""))

                st.subheader("Invoice Summary")
                gst_amount = st.text_input("GST Amount", data.get("GSTAmount", ""))
                total_amount = st.text_input("Total Amount After Tax", data.get("TotalAmountAfterTax", ""))

                st.subheader("Line Items (Editable)")
                corrected_items = []
                for idx, item in enumerate(data.get("Items", [])):
                    with st.expander(f"Line Item #{idx + 1}"):
                        item_number = st.text_input(f"Item #{idx + 1} - Number", item.get("ItemNumber", ""), key=f"item_num_{idx}")
                        item_desc = st.text_input(f"Item #{idx + 1} - Description", item.get("ItemDescription", ""), key=f"item_desc_{idx}")
                        item_qty = st.text_input(f"Item #{idx + 1} - Quantity", item.get("ItemQuantity", ""), key=f"item_qty_{idx}")
                        item_rate = st.text_input(f"Item #{idx + 1} - Rate", item.get("ItemRate", ""), key=f"item_rate_{idx}")
                        item_amt = st.text_input(f"Item #{idx + 1} - Amount", item.get("ItemAmount", ""), key=f"item_amt_{idx}")
                        item_hsn = st.text_input(f"Item #{idx + 1} - HSN/SAC", item.get("HSN/SAC", ""), key=f"item_hsn_{idx}")
                        item_gst_pct = st.text_input(f"Item #{idx + 1} - GST%", item.get("GST%", ""), key=f"item_gst_pct_{idx}")

                        corrected_items.append({
                            "ItemNumber": item_number,
                            "ItemDescription": item_desc,
                            "ItemQuantity": item_qty,
                            "ItemRate": item_rate,
                            "ItemAmount": item_amt,
                            "HSN/SAC": item_hsn,
                            "GST%": item_gst_pct,
                            "Total GST Amount": gst_amount,
                            "Total Amount After Tax": total_amount,
                            "Invoice Name": uploaded_file.name,
                            "InvoiceNumber": invoice_number,
                            "InvoiceDate": invoice_date,
                            "PO_Number": po_number,
                            "Buyer Date": buyer_date,
                            "Seller Name": seller_name,
                            "Seller Address": seller_address,
                            "Seller GST": seller_gst,
                            "Buyer Name": buyer_name,
                            "Buyer Address": buyer_address,
                            "Buyer GST": buyer_gst,
                        })

                if st.button("Submit and Save to Excel"):
                    df = pd.DataFrame(corrected_items)
                    if os.path.exists(output_excel_path):
                        existing_df = pd.read_excel(output_excel_path)
                        df = pd.concat([existing_df, df], ignore_index=True)
                    df.to_excel(output_excel_path, index=False)
                    st.success(f"Appended data to: `{output_excel_path}`")

            with col2:
                st.subheader("Scrollable Invoice Preview")
                show_pdf_scrollable(file_path)

        except Exception as e:
            st.error(f"Error: {e}")
