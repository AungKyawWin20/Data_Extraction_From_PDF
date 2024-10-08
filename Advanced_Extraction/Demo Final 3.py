import pdfplumber
from openpyxl import Workbook
from pathlib import Path

# Function to read section headers from a text file
def read_section_headers(file_path, encoding='utf-8'):
    with open(file_path, 'r', encoding=encoding) as file:
        return [line.strip() for line in file if line.strip()]

# Function to extract text and tables from a PDF page
def extract_text_and_tables(page):
    text = page.extract_text()
    tables = page.extract_tables()
    return text, tables

# Function to process a single page and return any tables found for the given headers
def process_page(page, section_headers):
    text, tables = extract_text_and_tables(page)
    if text:
        matching_headers = [header for header in section_headers if header in text]
        if matching_headers:
            return {header: tables for header in matching_headers if tables}
    return {}

# Function to extract tables associated with section headers from a PDF
def extract_tables_from_pdf(pdf_path, section_headers):
    tables_by_header = {header: [] for header in section_headers}
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            page_tables = process_page(page, section_headers)
            for header, tables in page_tables.items():
                tables_by_header[header].extend(tables)
    return tables_by_header

# Function to clean a table row by converting cells to strings and handling special cases
def clean_table_row(row):
    return [str(cell) if cell is not None and cell != -1 else '' for cell in row]

# Function to write tables to an Excel sheet, organized by section headers
def write_tables_to_excel(tables_by_header, ws):
    for header, tables in tables_by_header.items():
        if not tables:
            continue

        # Write section header
        ws.append([header])

        # Write each table
        for table in tables:
            for row in table:
                cleaned_row = clean_table_row(row)
                if any(cleaned_row):  # Avoid appending completely empty rows
                    ws.append(cleaned_row)
            ws.append([])  # Separate tables with an empty row

        ws.append([])  # Separate sections with an additional empty row

# Function to create and save the Excel workbook
def save_tables_to_excel(tables_by_header, excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Tables"
    write_tables_to_excel(tables_by_header, ws)
    wb.save(excel_path)

# Main function to orchestrate the extraction and saving of tables
def main():
    # Path to the text file with section headers
    headers_file_path = Path(r"Header_Files.txt").as_posix()
    # Path to PDF file
    pdf_path = Path(r"PDF_File.pdf").as_posix()
    # Path to the output Excel file
    excel_path = Path(r"Extracted Tables.xlsx").as_posix()

    section_headers = read_section_headers(headers_file_path)
    tables_by_header = extract_tables_from_pdf(pdf_path, section_headers)
    save_tables_to_excel(tables_by_header, excel_path)

    print(f"Tables have been successfully exported to {excel_path}")

if __name__ == "__main__":
    main()
