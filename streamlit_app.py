import streamlit as st
import json
import re
from io import BytesIO
import fitz  # PyMuPDF
from google.cloud import vision
from google.oauth2 import service_account
import openai
import openpyxl
import pdfplumber

# Load credentials
openai.api_key = st.secrets["OPENAI"]["OPENAI_API_KEY"]
creds_dict = json.loads(st.secrets["GCP"]["gcp_credentials"])
credentials = service_account.Credentials.from_service_account_info(creds_dict)
vision_client = vision.ImageAnnotatorClient(credentials=credentials)

# Upload CIM file (PDF or image)
st.title("📊 CIM to LBO Model Automation")
uploaded_cim = st.file_uploader("📄 Upload CIM (PDF or Image)", type=["pdf", "png", "jpg", "jpeg"])
uploaded_excel = st.file_uploader("📊 Upload Excel LBO Template", type=["xlsx"])

# Extract text from PDF/image
def extract_text(file):
    file_bytes = file.read()
    if file.type == "application/pdf":
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        is_digital = any(page.get_text() for page in doc)
        if is_digital:
            with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                return "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
        else:
            text = ""
            for page in doc:
                image_bytes = page.get_pixmap(dpi=300).tobytes("png")
                image = vision.Image(content=image_bytes)
                response = vision_client.document_text_detection(image=image)
                text += response.full_text_annotation.text + "\n"
            return text
    else:
        image = vision.Image(content=file_bytes)
        response = vision_client.document_text_detection(image=image)
        return response.full_text_annotation.text

# Build GPT prompt
def build_ai_prompt(text):
    return f"""
You are analyzing OCR output from a Confidential Information Memorandum (CIM) for an LBO model.

Extract these financials:

1. **Revenue** - Three most recent actual years + Six forward-looking years
2. **EBITDA** - Same structure as Revenue
3. **CapEx** 
4. **Acquisition Count** - Count of planned acquisitions per projected year

### Year Extraction Instructions:
- Identify all years tied to hardcoded financial values
- Sort chronologically and use:
  - Three earliest years for Actuals 
  - Next six years for forward-looking values 
  - If there is no info, make sure that when mapped to excel the cell becomes empty


### Candidate Metric Handling:
If multiple versions of a metric exist, ex: different ebitda or capex 

Group metric variants and wait for user input if multiple appear.

Text:
{text}
"""

# Begin processing
if uploaded_cim and uploaded_excel:
    st.success("✅ CIM and Excel uploaded")

    with st.spinner("📖 Extracting text..."):
        raw_text = extract_text(uploaded_cim)

    st.subheader("🔍 Extracted CIM Text")
    with st.expander("View Text"):
        st.text(raw_text)

    prompt = build_ai_prompt(raw_text)
    messages = [
        {"role": "system", "content": "You are a helpful assistant that ONLY outputs valid JSON."},
        {"role": "user", "content": prompt}
    ]

    with st.spinner("🤖 Extracting financials with GPT-4..."):
        response = openai.chat.completions.create(
            model="gpt-4",
            messages=messages,
            temperature=0
        )
        response_text = response.choices[0].message.content
        cim_data = json.loads(re.sub(r"```(json)?", "", response_text).strip())

    st.subheader("📈 Extracted Financial Data")
    st.json(cim_data)

    # Load Excel and extract metadata
    workbook = openpyxl.load_workbook(uploaded_excel)
    sheets = {name: workbook[name] for name in workbook.sheetnames}
    metadata = {}
    for name, sheet in sheets.items():
        formulas = []
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas.append((cell.coordinate, cell.value))
        metadata[name] = {
            "max_row": sheet.max_row,
            "max_col": sheet.max_column,
            "formulas": formulas
        }

    # Ask GPT to map values
    messages = [
        {"role": "system", "content": "You are a helpful assistant for populating Excel LBO models."},
        {"role": "user", "content": f"Excel metadata: {json.dumps(metadata)[:2000]}"},
        {"role": "user", "content": f"CIM extracted financials: {json.dumps(cim_data)[:2000]}"},
        {"role": "user", "content": "Map Revenue, EBITDA, CapEx, and Acquisitions to Excel cells. If multiple EBITDA options exist, ask me."}
    ]

    with st.spinner("🤖 Determining where to write data in Excel..."):
        response = openai.chat.completions.create(
            model="gpt-4",
            messages=messages,
            temperature=0
        )
        gpt_mapping = response.choices[0].message.content
        st.subheader("📌 GPT Mapping Suggestions")
        st.code(gpt_mapping)

    # Allow manual download of the unmodified file for now
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    st.download_button("📥 Download Excel", data=output, file_name="mapped_lbo_model.xlsx")
