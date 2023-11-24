from pypdf import PdfReader
import camelot
import shutil
import os
from collections import defaultdict


def is_selectable_text(pdf_path):
    print("-------------- Checking if text is selectable --------------")
    # Check if the PDF contains selectable text using PyPDF2
    try:
        with open(pdf_path, 'rb') as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                if text.strip():
                    return True
        return False
    except Exception as e:
        if str(e) == "Cannot read an empty file":
            print(f"Empty file encountered: {pdf_path}")
            return False
        else:
            print(f"Error processing {pdf_path}: {e}")
            return False


def does_not_contain_service_efr(pdf_path):
    # Check if the PDF does not contain "SERVICE EFR" or "HOPITAL FOCH" in any of its pages
    print("-------------- Checking if pdf in the word format --------------")
    try:
        with open(pdf_path, 'rb') as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                if "SERVICE EFR" in text.upper() or "HOPITAL FOCH"  in text.upper():  # Case-insensitive check
                    return False
        return True
    except Exception as e:
        if str(e) == "Cannot read an empty file":
            print(f"Empty file encountered: {pdf_path}")
            return False
        else:
            print(f"Error processing {pdf_path}: {e}")
            return False


def extract_tables(pdf_path):
    print("-------------- Extracting with camelot --------------")
    try:
        if is_selectable_text(pdf_path) and does_not_contain_service_efr(pdf_path):
            # Use Camelot to extract tables from PDF with selectable text and no "SERVICE EFR" or "HOPITAL FOCH"
            tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')
            return tables
    except Exception as e:
        if str(e) == "Cannot read an empty file":
            print(f"Empty file encountered: {pdf_path}")
            return False
        else:
            print(f"Error processing {pdf_path}: {e}")
            return False



def get_id_and_date_from_filename(filename):
    print("-------------- Formatting IPP  --------------")
    try:
        parts = filename.split('_')
        if len(parts) >= 2:
            return parts[0], parts[1]
        return None, None
    except Exception as e:
        if str(e) == "Cannot read an empty file":
            print(f"Empty file encountered: {filename}")
            return False
        else:
            print(f"Error processing {filename}: {e}")
            return False



def copy_pdfs_with_criteria(input_folder, output_folder):
    latest_files = defaultdict(lambda: ('', '0000-00-00'))

    for root, _, files in os.walk(input_folder):
        print(f"Currently scanning: {root}")
        for filename in files:
            print(f"Checking file: {filename}")
            if filename.endswith('.pdf'):
                pdf_path = os.path.join(root, filename)
                selectable = is_selectable_text(pdf_path)
                no_service_efr = does_not_contain_service_efr(pdf_path)
                print(f"Selectable: {selectable}, No Service EFR: {no_service_efr}")
                if selectable and no_service_efr:
                    id, date = get_id_and_date_from_filename(filename)
                    print(f"ID: {id}, Date: {date}, Latest Date: {latest_files[id][1]}")
                    if date > latest_files[id][1]:
                        latest_files[id] = (pdf_path, date)
                        destination_path = os.path.join(output_folder, filename)
                        print(f"Attempting to copy {pdf_path} to {destination_path}")
                        try:
                            shutil.copy(pdf_path, destination_path)
                            print(f"Successfully copied {filename}")
                        except Exception as e:
                            if str(e) == "Cannot read an empty file":
                                print(f"Empty file encountered: {filename}")
                                return False
                            else:
                                print(f"Error processing {filename}: {e}")
                                return False

def main():
    input_folder = "C:/Users/benysar/Desktop/LUTECE/extract_easily"
    output_folder = "C:/Users/benysar/Desktop/Github/OCR_EFR/QuickScanEFR/pdf_TextMachina/"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    copy_pdfs_with_criteria(input_folder, output_folder)


if __name__ == "__main__":
    main()
