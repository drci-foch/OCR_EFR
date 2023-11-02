import PyPDF2
import camelot
import shutil
import os
from collections import defaultdict


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
    # Check if the PDF does not contain "SERVICE EFR" or "HOPITAL FOCH" in any of its pages
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            if "SERVICE EFR" in text.upper() or "HOPITAL FOCH" in text.upper():  # Case-insensitive check
                return False
    return True


def extract_tables(pdf_path):
    if is_selectable_text(pdf_path) and does_not_contain_service_efr(pdf_path):
        # Use Camelot to extract tables from PDF with selectable text and no "SERVICE EFR" or "HOPITAL FOCH"
        tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')
        return tables


def get_id_and_date_from_filename(filename):
    # Assuming filenames are in the format "IPP_date_other-id-for-specific-document.pdf"
    parts = filename.split('_')
    if len(parts) >= 2:
        return parts[0], parts[1]
    return None, None


def copy_pdfs_with_criteria(input_folder, output_folder):
    # Default date to older one
    latest_files = defaultdict(lambda: ('', '0000-00-00'))

    for root, _, files in os.walk(input_folder):
        for filename in files:
            if filename.endswith('.pdf'):
                pdf_path = os.path.join(root, filename)
                if is_selectable_text(pdf_path) and does_not_contain_service_efr(pdf_path):
                    id, date = get_id_and_date_from_filename(filename)
                    # If this file is newer than the one stored
                    if date > latest_files[id][1]:
                        latest_files[id] = (filename, date)

    for id, (filename, date) in latest_files.items():
        shutil.copy(os.path.join(input_folder, filename),
                    os.path.join(output_folder, filename))


def main():
    input_folder = "../pdf_TextMachina"
    output_folder = "../pdf_selection"

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    copy_pdfs_with_criteria(input_folder, output_folder)


if __name__ == "__main__":
    main()
