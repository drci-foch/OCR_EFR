import PyPDF2
import camelot
import shutil
import os

## TODO : Keep only last one in date for each IPP

def is_selectable_text(pdf_path):
    # Check if the PDF contains selectable text using PyPDF2
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            if text.strip():
                return True
    return False

def does_not_contain_service_efr(pdf_path):
    # Check if the PDF does not contain "SERVICE EFR" in any of its pages
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            if "SERVICE EFR" in text.upper():  # Case-insensitive check
                return False
    return True

def extract_tables(pdf_path):
    if is_selectable_text(pdf_path) and does_not_contain_service_efr(pdf_path):
        # Use Camelot to extract tables from PDF with selectable text and no "SERVICE EFR"
        tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')
        return tables

def copy_pdfs_with_criteria(input_folder, output_folder):
    for root, _, files in os.walk(input_folder):
        for filename in files:
            if filename.endswith('.pdf'):
                pdf_path = os.path.join(root, filename)
                if is_selectable_text(pdf_path) and does_not_contain_service_efr(pdf_path):
                    shutil.copy(pdf_path, os.path.join(output_folder, filename))

def main():
    input_folder = "../pdf_TextMachina"  # Replace with the source folder path
    output_folder = "../pdf_selection"  # Replace with the destination folder path

    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    copy_pdfs_with_criteria(input_folder, output_folder)

if __name__ == "__main__":
    main()



