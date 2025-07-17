# Mini-Daloopa Prototype
# Goal: Upload a CIM-style PDF, extract financials using GPT, and populate an Excel LBO template.
import streamlit as st
import openai
import fitz  # PyMuPDF
import json
import openpyxl
from openpyxl.utils import get_column_letter
from io import BytesIO
import re

# ---- SETTINGS ----
openai.api_key = st.secrets["OPENAI"]["OPENAI_API_KEY"]  # Add in Streamlit secrets or replace directly
TEMPLATE_PATH = "TJC Practice Simple Model New (7) (2).xlsx"  # Replace with your Excel LBO model

# ---- UTILS ----
def extract_text_from_pdf(uploaded_file):
    pdf = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    full_text = ""
    for page in pdf:
        full_text += page.get_text()
    return full_text

def clean_json_response(response_text):
    """Clean and extract JSON from GPT response"""
    # Remove markdown code blocks if present
    response_text = re.sub(r'```json\n?', '', response_text)
    response_text = re.sub(r'```\n?', '', response_text)
    
    # Find JSON-like content between curly braces
    json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
    if json_match:
        return json_match.group(0)
    
    return response_text.strip()

def gpt_extract_financials(raw_text):
    prompt = f"""
You are a financial analyst. Extract the following key metrics from this CIM text:
- Revenue for each year (e.g., 2021, 2022, 2023)
- EBITDA for each year
- CapEx for each year

Return ONLY a valid JSON object with keys like Revenue_2021, EBITDA_2021, CapEx_2021, etc.
Use numeric values (not strings) for the financial figures.
If a value is not found, use null.

Example format:
{{
    "Revenue_2021": 100000000,
    "Revenue_2022": 120000000,
    "Revenue_2023": 140000000,
    "EBITDA_2021": 25000000,
    "EBITDA_2022": 30000000,
    "EBITDA_2023": 35000000,
    "CapEx_2021": 5000000,
    "CapEx_2022": 6000000,
    "CapEx_2023": 7000000
}}

Text:
{raw_text[:4000]}  # Only send first 4000 characters for demo
"""
    
    try:
        response = openai.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1  # Lower temperature for more consistent output
        )
        content = response.choices[0].message.content
        
        # Clean and parse JSON
        cleaned_content = clean_json_response(content)
        return json.loads(cleaned_content)
        
    except json.JSONDecodeError as e:
        st.error(f"JSON parsing error: {e}")
        st.error(f"GPT Response: {content}")
        return {}
    except Exception as e:
        st.error(f"Error calling GPT: {e}")
        return {}

def fill_excel_template(extracted_data):
    try:
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb["Model"]  # Assume data goes in 'Model' sheet
        
        # Map keys to named ranges (you can also hardcode cell names here)
        mapped_count = 0
        for key, value in extracted_data.items():
            try:
                # Try to find named range
                if key in wb.defined_names:
                    cell = wb.defined_names[key].destinations
                    for title, coord in cell:
                        target_ws = wb[title]
                        target_ws[coord] = value
                        mapped_count += 1
                else:
                    # If no named range, you could add hardcoded cell mappings here
                    st.warning(f"No named range found for: {key}")
            except Exception as e:
                st.warning(f"Could not map {key}: {e}")
        
        st.success(f"Successfully mapped {mapped_count} values to Excel template")
        
        output = BytesIO()
        wb.save(output)
        return output
        
    except FileNotFoundError:
        st.error(f"Excel template not found: {TEMPLATE_PATH}")
        return None
    except Exception as e:
        st.error(f"Error processing Excel template: {e}")
        return None

# ---- STREAMLIT APP ----
st.title("📊 Mini-Daloopa: Auto-fill Excel LBO Template")

st.markdown("""
This tool extracts financial data from CIM PDFs and populates an Excel LBO template.
""")

uploaded_file = st.file_uploader("Upload CIM PDF", type="pdf")

if uploaded_file:
    with st.spinner("Extracting text from PDF..."):
        raw_text = extract_text_from_pdf(uploaded_file)
    
    st.subheader("Extracted Text Preview")
    st.text_area("Preview (first 1000 characters)", raw_text[:1000], height=200)
    
    if st.button("Extract Financials & Fill Excel"):
        with st.spinner("Analyzing document with GPT..."):
            extracted_data = gpt_extract_financials(raw_text)
        
        if extracted_data:
            st.subheader("Extracted Financial Data")
            st.json(extracted_data)
            
            with st.spinner("Populating Excel template..."):
                output_excel = fill_excel_template(extracted_data)
            
            if output_excel:
                st.success("Excel template populated successfully!")
                st.download_button(
                    label="Download Populated Model",
                    data=output_excel.getvalue(),
                    file_name="lbo_filled.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("Failed to extract financial data. Please check the PDF content and try again.")

# ---- DEBUGGING SECTION ----
with st.expander("Debug Information"):
    st.write("Template path:", TEMPLATE_PATH)
    if uploaded_file:
        st.write("Text length:", len(raw_text))
        st.write("First 500 characters:", raw_text[:500])
