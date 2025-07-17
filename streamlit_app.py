# Mini-Daloopa Prototype
# Goal: Upload a CIM-style PDF, extract financials using GPT, and populate an Excel LBO template.

import streamlit as st
import openai
import fitz  # PyMuPDF
import json
import openpyxl
from openpyxl.utils import get_column_letter
from io import BytesIO

# ---- SETTINGS ----
openai.api_key = st.secrets["OPENAI_API_KEY"]  # Add in Streamlit secrets or replace directly
TEMPLATE_PATH = "lbo_template.xlsx"  # Replace with your Excel LBO model

# ---- UTILS ----
def extract_text_from_pdf(uploaded_file):
    pdf = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    full_text = ""
    for page in pdf:
        full_text += page.get_text()
    return full_text

def gpt_extract_financials(raw_text):
    prompt = f"""
You are a financial analyst. Extract the following key metrics from this CIM text:
- Revenue for each year (e.g., 2021, 2022, 2023)
- EBITDA for each year
- CapEx for each year
Return as JSON with keys like Revenue_2021, EBITDA_2021, etc.

Text:
{raw_text[:4000]}  # Only send first 4000 tokens for demo
"""
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )
    content = response.choices[0].message.content
    return json.loads(content)

def fill_excel_template(extracted_data):
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb["Model"]  # Assume data goes in 'Model' sheet

    # Map keys to named ranges (you can also hardcode cell names here)
    for key, value in extracted_data.items():
        try:
            cell = wb.defined_names[key].destinations
            for title, coord in cell:
                target_ws = wb[title]
                target_ws[coord] = value
        except:
            st.warning(f"Could not map: {key}")

    output = BytesIO()
    wb.save(output)
    return output

# ---- STREAMLIT APP ----
st.title("📊 Mini-Daloopa: Auto-fill Excel LBO Template")
uploaded_file = st.file_uploader("Upload CIM PDF", type="pdf")

if uploaded_file:
    raw_text = extract_text_from_pdf(uploaded_file)
    st.text_area("Extracted Text Preview", raw_text[:1000], height=200)

    if st.button("Extract Financials & Fill Excel"):
        try:
            extracted_data = gpt_extract_financials(raw_text)
            st.json(extracted_data)

            output_excel = fill_excel_template(extracted_data)
            st.download_button("Download Populated Model", data=output_excel.getvalue(), file_name="lbo_filled.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")

