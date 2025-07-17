import openpyxl
import json
import openai

class ExcelLBOAssistant:
    def __init__(self, template_path):
        self.template_path = template_path
        self.workbook = openpyxl.load_workbook(self.template_path)
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

    def get_cell(self, sheet_name, cell):
        return str(self.sheets[sheet_name][cell].value)

    def set_cell(self, sheet_name, cell, value):
        self.sheets[sheet_name][cell].value = value
        return True

    def get_errors(self, sheet_name):
        errors = []
        for row in self.sheets[sheet_name].iter_rows():
            for cell in row:
                if cell.data_type == 'e':
                    errors.append((cell.coordinate, cell.value))
        return errors

    def save(self, output_path):
        self.workbook.save(output_path)

# Sample GPT integration
 openai.api_key = st.secrets["OPENAI"]["OPENAI_API_KEY"]

assistant = ExcelLBOAssistant("TJC Practice Simple Model New (7).xlsx")
metadata = assistant.get_metadata()

messages = [
    {"role": "system", "content": "You are an LBO financial model assistant."},
    {"role": "user", "content": f"This is the metadata of the Excel model: {json.dumps(metadata)[:1000]}..."},
    {"role": "user", "content": "Check if there are any formula errors or unfilled EBITDA cells. Ask me which version of EBITDA to use if multiple exist."}
]

response = openai.ChatCompletion.create(
    model="gpt-4",
    messages=messages,
    temperature=0
)

print(response.choices[0].message.content)
