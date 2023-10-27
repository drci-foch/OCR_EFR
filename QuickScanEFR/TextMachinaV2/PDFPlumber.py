import camelot
import pandas as pd
import os 
# Refactoring the PDFPlumber.py module

class PDFProcessorRefactored:
    def __init__(self, file_path=None):
        self.file_path = file_path

    def process_pdf(self, file_path):
        """Extract tables from a PDF file and save them as an Excel file."""
        try:
            tables = camelot.read_pdf(file_path, pages='all', flavor='stream')
            with pd.ExcelWriter(file_path.replace(".pdf", ".xlsx")) as writer:
                for idx, table in enumerate(tables):
                    table.df.to_excel(writer, sheet_name=f"Table_{idx}", index=False)
            print(f"Processed {file_path}")
        except Exception as e:
            print(f"Error processing {file_path}. Error: {str(e)}")

    def process_directory(self, directory_path):
        """Process all PDF files in the specified directory."""
        for filename in os.listdir(directory_path):
            if filename.endswith(".pdf"):
                file_path = os.path.join(directory_path, filename)
                self.process_pdf(file_path)


