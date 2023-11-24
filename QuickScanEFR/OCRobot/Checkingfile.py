from pypdf import PdfReader
from pypdf.errors import EmptyFileError  # Add this line to import the exception
from pypdf.errors import PdfStreamError
import shutil
import os
from collections import defaultdict


def is_not_selectable_text(pdf_path):
    try:
        with open(pdf_path, 'rb') as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                if text.strip():
                    return False
    except (PdfStreamError, EmptyFileError) as e:
        print(f"Error processing file {pdf_path}: {e}")
    return True


def get_id_and_date_from_filename(filename):
    # Assuming filenames are in the format "IPP_date_other-id-for-specific-document.pdf"
    parts = filename.split('_')
    if len(parts) >= 2:
        return parts[0], parts[1]
    return None, None


def copy_pdfs_with_criteria(input_folder, output_folder):
    # Get the list of all PDF files in the input folder
    pdf_files = [os.path.join(root, f) for root, _, files in os.walk(input_folder) for f in files if f.endswith('.pdf')]
    total_files = len(pdf_files)
    processed_files = 0

    if total_files == 0:
        print("No PDF files found in the directory.")
        return

    # Default date to older one
    latest_files = defaultdict(lambda: ('', '0000-00-00'))

    for pdf_path in pdf_files:
        print(f"Checking file: {pdf_path}")  # Debug print to check file detection
        try:
            if is_not_selectable_text(pdf_path):
                filename = os.path.basename(pdf_path)
                id, date = get_id_and_date_from_filename(filename)
                # If this file is newer than the one stored
                if date > latest_files[id][1]:
                    latest_files[id] = (filename, date)
                    # Copy the file immediately if it's the latest
                    source_path = pdf_path
                    destination_path = os.path.join(output_folder, filename)
                    print(f"Copying file {source_path} to {destination_path}")  # Debug print for copying
                    shutil.copy(source_path, destination_path)
        except Exception as e:
            print(f"Error processing {pdf_path}: {e}")

        processed_files += 1
        progress = (processed_files / total_files) * 100
        print(f"Processing... {processed_files}/{total_files} ({progress:.2f}%)")




def main():
    input_folder = "C:/Users/benysar/Desktop/LUTECE/extract_easily/"
    output_folder = "../pdf_OCRobot"

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    copy_pdfs_with_criteria(input_folder, output_folder)


if __name__ == "__main__":
    main()
