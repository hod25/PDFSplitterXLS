import os
import pandas as pd
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter
import re

def extract_date_from_filename(filename):
    # Define a regular expression pattern to extract date from filename
    pattern = r'\b(\d+\.\d+)\b'  # Assumes the pattern is digits followed by a dot and more digits
    
    # Use regular expression to search for date pattern in filename
    match = re.search(pattern, filename)
    
    if match:
        return match.group(1)  # Return the matched date string
    else:
        return None  # Return None if no date found in filename

def split_pdf():
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Find the PDF file in the script directory
    pdf_files = [f for f in os.listdir(script_dir) if f.endswith('.pdf')]
    if not pdf_files:
        print(f"Error: No PDF file found in the directory {script_dir}.")
        return
    
    input_pdf_filename = pdf_files[0]  # Get the name of the first PDF file found
    input_pdf_path = os.path.join(script_dir, input_pdf_filename)
    
    # Extract date from the original PDF file name
    date_from_filename = extract_date_from_filename(input_pdf_filename)
    
    if date_from_filename:
        print(f"Date extracted from original PDF filename: {date_from_filename}")
    else:
        print("Error: No date found in the original PDF filename.")
        return
    
    excel_file_path = os.path.join(script_dir, 'names.xlsx')
    
    # Check if the Excel file exists
    if not os.path.exists(excel_file_path):
        print(f"Error: Excel file 'names.xlsx' not found in the directory {script_dir}.")
        return
    
    # Read the Excel file
    try:
        df = pd.read_excel(excel_file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    # Check if the 'Filename' column exists
    if 'Filename' not in df.columns:
        print(f"Error: 'Filename' column not found in the Excel file.")
        return
    
    # Read the PDF file
    try:
        pdf_reader = PdfReader(input_pdf_path)
    except Exception as e:
        print(f"Error reading PDF file: {e}")
        return
    
    # Specify the path for the new folder
    pdf_name = os.path.splitext(input_pdf_filename)[0]
    folder_path = Path(script_dir) / pdf_name
    folder_path.mkdir(parents=True, exist_ok=True)
    
    print(f"Folder for split PDF pages created at: {folder_path}")
    
    # Split the PDF into individual pages
    for page_num in range(len(pdf_reader.pages)):
        pdf_writer = PdfWriter()
        pdf_writer.add_page(pdf_reader.pages[page_num])
        
        # Get the filename for the current page from the Excel file
        try:
            output_filename = df.loc[page_num, 'Filename']
            print(f"PDF Name: {output_filename}")  # Print the PDF name
        except KeyError:
            output_filename = f'page_{page_num + 1}.pdf'
        
        # Append the date from the original PDF filename to the output filename
        output_filename_with_date = f'{date_from_filename} {output_filename}'
        
        output_pdf_path = folder_path / output_filename_with_date
        try:
            with open(output_pdf_path, 'wb') as output_pdf_file:
                pdf_writer.write(output_pdf_file)
            print(f'Page {page_num + 1} saved as {output_pdf_path}')
        except Exception as e:
            print(f"Error writing PDF file {output_pdf_path}: {e}")

if __name__ == '__main__':
    split_pdf()

# !/usr/bin/env python3
# Path: python split_pdf.py