import tkinter as tk
from tkinter import filedialog
from PyPDFForm import PdfWrapper, FormWrapper
from jinja2 import Template
from textwrap import dedent
import yaml, openpyxl, os

def select_pdf_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )
    return file_path

def select_yaml_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select a map yaml file",
        filetypes=[("YAML files", "*.yml *.yaml")]
    )
    return file_path

def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select a excel file",
        filetypes=[("XLS files", "*.xls *.xlsx *.xlsm")]
    )
    return file_path

def gen_tagged(pdf_path):
    try:
        pdf_form = PdfWrapper(pdf_path)
        preview_stream = pdf_form.preview
        tagged_path = pdf_path.replace('.pdf', '_fields.pdf')
        with open(tagged_path, "wb+") as output:
            output.write(preview_stream)
    except Exception as e:
        print(f"Error generating tagged pdf: {e}")
        return None

def gen_map(pdf_path):
    try:
        map_template_str = dedent("""\
        fields:
          {% for field in fields -%}
          '{{ field }}':
            sheet: ''
            cell: ''
          {% endfor %}
        """)
        map_template = Template(map_template_str)
        pdf_form = PdfWrapper(pdf_path)
        fields = pdf_form.schema["properties"]
        string_fields = fields #[k for k, v in fields.items() if v['type'] == 'string']
        rendered_map = map_template.render({"fields": string_fields})
        map_path = pdf_path.replace('.pdf', '_map.yml')
        with open(map_path, "w+") as output:
            output.write(rendered_map)
    except Exception as e:
        print(f"Error generating yaml map: {e}")
        return None

def load_map(yaml_path):
    try:
        with open(yaml_path, 'r') as f:
            map = yaml.safe_load(f).get("fields", {})
        map_nonempty = {k: v for k, v in map.items() if (v['sheet'] != '' and v['cell'] != '')}
        return map_nonempty
    except Exception as e:
        print(f"Error loading yaml map: {e}")
        return None

def extract_values(excel_path, map):
    workbook = openpyxl.load_workbook(excel_path, data_only=True)
    extracted_values = {}
    for field, spec in map.items():
        sheet = spec.get("sheet")
        cell = spec.get("cell")
        if sheet not in workbook.sheetnames:
            print(f"Warning: Sheet '{sheet}' not found in workbook.")
            continue
        sheet = workbook[sheet]
        value = sheet[cell].value
        extracted_values[field] = str(value)
    return extracted_values

def fill_form(pdf_path, extracted_values):
    try:
        pdf_form = FormWrapper(pdf_path)
        filled = pdf_form.fill(extracted_values)
        filled_path = pdf_path.replace('.pdf', '_filled.pdf')
        with open(filled_path, "wb+") as output:
            output.write(filled.read())
    except Exception as e:
        print(f"Error filling pdf form: {e}")
        return None

def run_read_form():
    print("Select a PDF file with form fields.")
    pdf_path = select_pdf_file()

    if not pdf_path:
        print("No file selected.")
        return

    print(f"PDF selected: {pdf_path}")
    print(f"Generating tagged version")
    gen_tagged(pdf_path)
    gen_map(pdf_path)
    print(f"Done!")
    return

def run_fill_form():
    print("Select yaml map file.")
    map_path = select_yaml_file()
    map = load_map(map_path)
    print("Select excel file.")
    excel_path = select_excel_file()
    extracted_values = extract_values(excel_path, map)
    print("Select PDF file.")
    pdf_path = select_pdf_file()
    print("Filling Form.")
    fill_form(pdf_path, extracted_values)
    print(f"Done!")
    return

def main():
    print("Please choose an option:")
    print("1 - Extract map and field names from PDF form")
    print("2 - Fill form from map and Excel file")

    choice = input("Enter your choice (1 or 2): ").strip()

    if choice == '1':
        run_read_form()
    elif choice == '2':
        run_fill_form()
    else:
        print("Invalid choice. Please enter 1 or 2.")

