import pdfplumber
from openpyxl import Workbook

# Function to extract tables from the PDF except the first page
def extract_tables_from_pdf(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num in range(1, len(pdf.pages)):
            page = pdf.pages[page_num]
            tables.extend(page.extract_tables())
    return tables

# Function to export tables and checkboxes to an Excel file
def export_tables_to_excel(tables, checkboxes, excel_path):
    wb = Workbook()
    ws = wb.active
    # Write tables to Excel
    for table in tables:
        for row in table:
            ws.append(row)
        ws.append([])
    # Write checkbox information to a new sheet
    ws_checkboxes = wb.create_sheet("Checkboxes")
    ws_checkboxes.append(["Page", "x0", "y0", "x1", "y1", "Checked"])
    for checkbox in checkboxes:
        ws_checkboxes.append([
            checkbox["page"],
            checkbox["x0"],
            checkbox["y0"],
            checkbox["x1"],
            checkbox["y1"],
            checkbox["checked"]
        ])
    wb.save(excel_path)

# Function to extract checkboxes from the PDF
def extract_checkboxes_from_pdf(pdf_path):
    checkboxes = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num in range(1, len(pdf.pages)):
            page = pdf.pages[page_num]
            rects = page.rects
            for rect in rects:
                # Filter out small rectangles that are unlikely to be checkboxes
                if 30 < rect["width"] < 100 and 30 < rect["height"] < 100 and int(rect["width"]) == int(rect["height"]):
                    # Extract the box color
                    box_color = rect['non_stroking_color']

                    # Check if the color of the cross (if it exists) matches the box color
                    is_checked = is_cross_present(page, rect) and is_cross_color_matching(page, rect, box_color)

                    checkboxes.append({
                        "page": page_num + 1,
                        "x0": rect["x0"],
                        "y0": rect["y0"],
                        "x1": rect["x1"],
                        "y1": rect["y1"],
                        "checked": is_checked
                    })
    return checkboxes

def is_cross_present(page, rect):
    # Implement logic to check if a cross is present in the box
    # For simplicity, let's assume we have a way to determine this
    # This function needs to be defined according to the PDF's content structure
    return True

def is_cross_color_matching(page, rect, box_color):
    # Implement logic to check if the color of the cross matches the box color
    # This is a placeholder function and needs to be defined based on how colors are represented in the PDF
    return True

# Path to PDF file
pdf_path = "17364-725846.pdf"
# Path to the output Excel file
excel_path = "extracted_tables_checkboxes.xlsx"

# Extract tables from the PDF except the first page
tables = extract_tables_from_pdf(pdf_path)

# Extract checkboxes from the PDF
checkboxes = extract_checkboxes_from_pdf(pdf_path)

# Export the extracted tables and checkboxes to an Excel file
export_tables_to_excel(tables, checkboxes, excel_path)

print(f"Tables and checkboxes have been successfully exported to {excel_path}")
