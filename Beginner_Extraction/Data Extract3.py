import pdfplumber
from openpyxl import Workbook

def extract_tables_from_pdf(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num in range(1, len(pdf.pages)):
            page = pdf.pages[page_num]
            tables.extend(page.extract_tables())
    return tables

def export_tables_to_excel(tables, excel_path):
    wb = Workbook()
    ws = wb.active
    for table in tables:
        for row in table:
            ws.append(row)
        ws.append([])
    wb.save(excel_path)

# Path to PDF file
pdf_path = "17364-725846.pdf"
# Path to the output Excel file
excel_path = "extracted_tables100.xlsx"

# Extract tables from the PDF except the first page
tables = extract_tables_from_pdf(pdf_path)

# Export the extracted tables to an Excel file
export_tables_to_excel(tables, excel_path)

print(f"Tables have been successfully exported to {excel_path}")
