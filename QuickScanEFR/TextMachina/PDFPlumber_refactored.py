import os
import camelot
import pandas as pd
import threading

class PDFProcessor:
    def __init__(self, directory_path):
        self.directory_path = directory_path

    def process_pdf(self, pdf_path, output_path):
        tables = camelot.read_pdf(pdf_path, pages='all', strip_text='\n')
        with pd.ExcelWriter(output_path) as writer:
            for i, table in enumerate(tables):
                table.df.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)

    def process_directory(self):
        for filename in os.listdir(self.directory_path):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(self.directory_path, filename)
                output_path = os.path.join(self.directory_path.replace('../pdf_TextMachina', '../pdf_TextMachina/excel'), filename.replace('.pdf', '.xlsx'))

                # Create a thread for processing each PDF
                thread = threading.Thread(target=self.process_pdf, args=(pdf_path, output_path))
                thread.start()

if __name__ == "__main__":
    pipeline = PDFProcessor('../pdf_TextMachina')
    pipeline.process_directory()
