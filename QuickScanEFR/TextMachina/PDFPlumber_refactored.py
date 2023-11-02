import os
import camelot
import pandas as pd
import multiprocessing


class PDFProcessor:
    """
    A class to process PDFs in a given directory, extract tables, 
    and save them as Excel files.
    
    Attributes:
        directory_path (str): Path to the directory containing PDF files.
    """

    def __init__(self, directory_path):
        """
        Initializes the PDFProcessor with the given directory path.
        
        Args:
            directory_path (str): Path to the directory containing PDF files.
        """
        self.directory_path = directory_path

    def process_pdf(self, pdf_path, output_path):
        """
        Processes a single PDF, extracts tables from it, 
        and saves them in an Excel file.
        
        Args:
            pdf_path (str): Path to the PDF file.
            output_path (str): Path to save the resulting Excel file.
        """
        tables = camelot.read_pdf(pdf_path, pages='all', strip_text='\n')
        with pd.ExcelWriter(output_path) as writer:
            for i, table in enumerate(tables):
                table.df.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)

    def combine_data_horizontally(self, excel_path):
        """
        Combines sheets of an Excel file horizontally 
        based on the first column.
        
        Args:
            excel_path (str): Path to the Excel file.
        
        Returns:
            DataFrame: Combined data from all the sheets.
        """
        xls = pd.ExcelFile(excel_path)
        sheet_names = xls.sheet_names
        combined_data = pd.read_excel(xls, sheet_name=sheet_names[0])
        for sheet in sheet_names[1:]:
            temp_df = pd.read_excel(xls, sheet_name=sheet)
            temp_df = temp_df.drop(temp_df.columns[0], axis=1)
            combined_data = pd.concat([combined_data, temp_df], axis=1)
        return combined_data

    def process_directory(self):
        """
        Processes all PDF files in the directory 
        using multiprocessing for efficiency.
        """
        pool = multiprocessing.Pool(processes=multiprocessing.cpu_count())

        for filename in os.listdir(self.directory_path):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(self.directory_path, filename)
                output_path = os.path.join(self.directory_path.replace(
                    '../pdf_TextMachina', '../pdf_TextMachina/excel'), filename.replace('.pdf', '.xlsx'))

                pool.apply_async(self.process_and_combine,
                                 args=(pdf_path, output_path))

        pool.close()
        pool.join()

    def process_and_combine(self, pdf_path, output_path):
        """
        Processes a PDF and combines the extracted tables horizontally.
        
        Args:
            pdf_path (str): Path to the PDF file.
            output_path (str): Path to save the resulting Excel file.
        """
        self.process_pdf(pdf_path, output_path)
        combined_data = self.combine_data_horizontally(output_path)
        combined_data.to_excel(output_path, index=False)


if __name__ == "__main__":
    pipeline = PDFProcessor('../pdf_TextMachina')
    pipeline.process_directory()
