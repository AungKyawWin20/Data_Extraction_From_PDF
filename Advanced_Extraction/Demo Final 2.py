import pdfplumber
from openpyxl import Workbook
from pathlib import Path

# Function to read section headers from a text file
def read_section_headers(file_path, encoding='utf-8'):
    with open(file_path, 'r', encoding=encoding) as file:
        section_headers = file.read().splitlines()
    return section_headers


# Function to extract specific tables from a PDF
def extract_specific_tables(pdf_path, section_headers):
    tables_by_header = {header: [] for header in section_headers}
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            print(f"Processing page {page_num + 1}...")
            if text:
                for header in section_headers:
                    if header in text:
                        print(f"Found header '{header}' on page {page_num + 1}")  # Debugging log
                        tables_by_header[header].extend(page.extract_tables())
            else:
                print(f"No text found on page {page_num + 1}")  # Debugging log
    return tables_by_header


# Function to export tables to an Excel file
def export_tables_to_excel(tables_by_header, excel_path):
    wb = Workbook()
    wb.remove(wb.active)
    for header, tables in tables_by_header.items():
        ws = wb.create_sheet(title=header[:30])
        for table in tables:
            for row in table:
                ws.append(row)
            ws.append([])
    wb.save(excel_path)


# Main function
def main():
    # Path to the text file with section headers
    headers_file_path = Path(r"Header_Files.txt").as_posix()
    # Path to PDF file
    pdf_path = Path(r"PDF_File.pdf").as_posix()
    # Path to the output Excel file
    excel_path = Path(r"Extracted Tables.xlsx").as_posix()

    # Read section headers from the text file with specified encoding
    section_headers = read_section_headers(headers_file_path, encoding='utf-8')

    # Extract specific tables from the PDF
    tables_by_header = extract_specific_tables(pdf_path, section_headers)

    # Export the extracted tables to an Excel file
    export_tables_to_excel(tables_by_header, excel_path)

    print(f"Tables have been successfully exported to {excel_path}")


if __name__ == "__main__":
    main()
