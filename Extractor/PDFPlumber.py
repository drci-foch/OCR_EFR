import camelot
import pandas as pd
import os

def process_pdf(pdf_path):
    """
    Process a single PDF file and extract tables to save in an Excel file.
    """
    tables = camelot.read_pdf(pdf_path, pages='all', strip_text='\n')

    print(f"Processing {pdf_path}. Total tables extracted: {tables.n}")

    output_filename = os.path.splitext(os.path.basename(pdf_path))[0] + '.xlsx'
    output_filepath = os.path.join(os.path.dirname(pdf_path), output_filename)

    with pd.ExcelWriter(output_filepath) as writer:
        for i, table in enumerate(tables):
            df = table.df
            
            df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)

    print(f"Saved tables from {pdf_path} to {output_filepath}")

def main(directory_path):
    """
    Process all PDF files in the specified directory.
    """
    for filename in os.listdir(directory_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(directory_path, filename)
            process_pdf(pdf_path)

if __name__ == '__main__':
    directory_path = './pdf'
    main(directory_path)
