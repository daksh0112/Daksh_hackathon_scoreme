import fitz  
from openpyxl import Workbook
import re

def clean_text(text):
    """Remove colons and extra spaces from text."""
    text = text.replace(":", "").strip() 
    text = text.replace("-", "").strip()
    text = re.sub(r'\s{2,}', '\t', text) 
    return text

def extract_text_tables(pdf_path, output_excel):
    """Extracts structured text-based tables from a PDF and saves them to Excel."""
    doc = fitz.open(pdf_path)
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Extracted Tables"

    for page in doc:
        text = page.get_text("text")  
        lines = text.split("\n")  

        structured_data = []  

        for line in lines:
            line = clean_text(line)  
            row_data = line.split("\t")  

            if row_data:
                structured_data.append(row_data)

        
        for row in structured_data:
            sheet.append(row)

    wb.save(output_excel)
    print(f"Tables extracted and saved to {output_excel}")


pdf_file = "/content/test3 (1).pdf"  
excel_file = "/content/extracted_tables.xlsx"  

extract_text_tables(pdf_file, excel_file)