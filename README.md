# Extracting Tables from PDFs Using Python

## Overview

These Python scripts extract tables from PDF files based on section headers specified in a text file and export them to an Excel spreadsheet. Each script takes a slightly different approach to processing and storing tables.

## Prerequisites

Ensure you have the following Python libraries installed before running any of the scripts:
```bash
pip install pdfplumber openpyxl
```

## Types of Scripts

### Script 1: Extract Tables By Section Header
- Reads section headers from a text file.
- Iterates through the PDF pages and extracts tables if a header is found on the page.
- Stores tables in a dictionary categorized by headers.
- Exports tables to an Excel file, creating separate sheets for each header.

### Script 2: Optimized Extraction with Cleaning and Single-Sheet Output
- Reads section headers and filters tables based on the presence of headers.
- Extracts both text and tables from each page.
- Cleans tables by converting non-standard values.
- Stores extracted tables under each header and writes them into a single Excel sheet.

### Script 3: Extract All Tables Without Categorization
- Reads section headers but does not categorize tables under specific headers.
- Extracts all tables from the PDF if any header is found on a page.
- Writes extracted tables to an Excel file in a single sheet, separated by blank rows.

## Troubleshooting

Troubleshooting

- If the script does not extract tables correctly, check if:
  - The section headers match the text formatting in the PDF.
  - The PDF contains extractable tables (some PDFs use images instead of text-based tables).
  - You have the required Python libraries installed.

## Summary
This set of scripts is designed to facilitate automated extraction and analysis of structured data from PDFs using Python.
