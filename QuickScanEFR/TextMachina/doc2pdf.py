import win32com.client
import os
from pathlib import Path

def convert_word_to_pdf(doc_path, pdf_path=None):
    if pdf_path is None:
        pdf_path = str(Path(doc_path).with_suffix('.pdf'))
        
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        print(f"Opening {doc_path}")
        doc = word.Documents.Open(str(Path(doc_path).absolute()))
        print(f"Saving as PDF: {pdf_path}")
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()
        return pdf_path
    except Exception as e:
        print(f"Error in conversion: {str(e)}")
        if 'word' in locals():
            word.Quit()
        raise

def batch_convert_directory(directory):
    directory = Path(directory)
    print(f"Scanning directory: {directory}")
    
    # Find all .doc and .docx files
    word_files = list(directory.glob('*.doc')) + list(directory.glob('*.docx'))
    print(f"Found {len(word_files)} Word files")
    
    for doc_file in word_files:
        try:
            print(f"\nProcessing: {doc_file}")
            convert_word_to_pdf(str(doc_file))
            print("Conversion successful")
        except Exception as e:
            print(f"Failed to convert {doc_file}: {str(e)}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        directory = sys.argv[1]
        print(f"Starting conversion in directory: {directory}")
        batch_convert_directory(directory)
    else:
        print("Please provide a directory path")