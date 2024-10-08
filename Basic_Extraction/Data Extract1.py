#Importing libraries
import pdfplumber
import pandas as pd

#Open the file
file = "Summary-of-Opioid-Funds-to-Virginia-Localities-as-of-Jan-2023.pdf"

#Select the first page
with pdfplumber.open(file) as pdf:
    page = pdf.pages[0]
#Extract the tables
    tables = page.extract_tables()
    print(tables)

#Iterate over each table and save it in a seperate Excel file
    for i, table in enumerate(tables):
        df = pd.DataFrame(table[1:], columns = table[0])
        print(df)
        excel_file = f"Data_Table{i+1}.xlsx"

#Save it to your desired location
        df.to_excel(excel_file, index = False)

        print(f"Table {i+1} extracted and saved to {excel_file}")
