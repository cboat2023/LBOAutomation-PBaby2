import streamlit as st
import openai
import openpyxl
import json
from io import BytesIO

# Setup OpenAI client
client = openai.OpenAI(api_key=st.secrets["OPENAI"]["OPENAI_API_KEY"])

class ExcelLBOAssistant:
    def __init__(self, file_obj):
        self.workbook = openpyxl.load_workbook(file_obj)
        self.sheets = {name: self.workbook[name] for name in self.workbook.sheetnames}

    def get_metadata(self):
        metadata = {}
        for name, sheet in self.sheets.items():
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
        return metadata

    def get_errors(self):
        errors = []
        for name, sheet in self.sheets.items():
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == 'e':
                        errors.append((name, cell.coordinate, cell.value))
        return errors

    def save(self):
        output = BytesIO()
        self.workbook.save(output)
        output.seek(0)
        return output

# UI
st.title("📊 GPT-4 Excel LBO Assistant")
uploaded_file = st.file_uploader("📁 Upload your Excel LBO Model", type=["xlsx"])

if uploaded_file:
    st.success("✅ File uploaded successfully!")
    assistant = ExcelLBOAssistant(uploaded_file)
    metadata = assistant.get_metadata()

    with st.spinner("🤖 Asking GPT to analyze the Excel structure..."):
        messages = [
            {"role": "system", "content": "You are an LBO financial model assistant."},
            {"role": "user", "content": f"This is the metadata of the Excel model: {json.dumps(metadata)[:1000]}..."},
            {"role": "user", "content": "Check if there are any formula errors or unfilled EBITDA cells. Ask me which version of EBITDA to use if multiple exist."}
        ]

        try:
            response = client.chat.completions.create(
                model="gpt-4",
                messages=messages,
                temperature=0
            )
            gpt_reply = response.choices[0].message.content
            st.subheader("🧠 GPT Response")
            st.write(gpt_reply)
        except Exception as e:
            st.error(f"❌ GPT error: {e}")

    # Show Excel formula errors
    errors = assistant.get_errors()
    if errors:
        st.warning("⚠️ Formula errors found:")
        for sheet, cell, val in errors:
            st.write(f"- `{sheet}!{cell}` → {val}")
    else:
        st.success("✅ No formula errors detected.")

    # Download updated file
    excel_bytes = assistant.save()
    st.download_button(
        label="📥 Download Updated Excel Model",
        data=excel_bytes,
        file_name="updated_lbo_model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

