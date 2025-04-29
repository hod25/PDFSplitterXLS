# PDF Splitter with Named Pages #

This Python script splits a PDF file into individual pages and names each page according to an Excel file (`names.xlsx`). It also appends a date extracted from the original PDF filename to each output file name.

---

## Features

- Automatically detects the first PDF file in the directory
- Extracts a date (e.g., `12.2023`) from the PDF filename
- Reads a list of target filenames from an Excel file
- Splits the PDF into separate pages
- Saves each page as a new PDF with a name from the Excel file, prefixed with the extracted date
- Organizes all output files into a folder based on the original PDF name

## Requirements

- Python 3.6+
- `pandas`
- `PyPDF2`
- `openpyxl` (used by `pandas` to read Excel files)

Install the required packages using pip:

```bash
pip install pandas PyPDF2 openpyxl
