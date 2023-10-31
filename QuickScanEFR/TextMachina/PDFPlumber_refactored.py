import os
import camelot
import pandas as pd
import multiprocessing

class PDFProcessor:
    def __init__(self, directory_path):
        self.directory_path = directory_path

    def process_pdf(self, pdf_path, output_path):
        tables = camelot.read_pdf(pdf_path, pages='all', strip_text='\n')
        with pd.ExcelWriter(output_path) as writer:
            for i, table in enumerate(tables):
                table.df.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)

    def combine_data_horizontally(self, excel_path):
        xls = pd.ExcelFile(excel_path)
        sheet_names = xls.sheet_names
        combined_data = pd.read_excel(xls, sheet_name=sheet_names[0])
        for sheet in sheet_names[1:]:
            temp_df = pd.read_excel(xls, sheet_name=sheet)
            temp_df = temp_df.drop(temp_df.columns[0], axis=1)  # Drop the first column
            combined_data = pd.concat([combined_data, temp_df], axis=1)
        return combined_data

    def process_directory(self):
        # Create a multiprocessing pool
        pool = multiprocessing.Pool(processes=multiprocessing.cpu_count())

        for filename in os.listdir(self.directory_path):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(self.directory_path, filename)
                output_path = os.path.join(self.directory_path.replace('../pdf_TextMachina', '../pdf_TextMachina/excel'), filename.replace('.pdf', '.xlsx'))

                # Use the pool to process each PDF
                pool.apply_async(self.process_pdf, args=(pdf_path, output_path))

        # Close the pool and wait for all processes to finish
        pool.close()
        pool.join()

        # Combine data from all processed Excel files horizontally
        combined_data = self.combine_data_horizontally(output_path)

        # Save the combined data to a new Excel file
        combined_output_path = os.path.join(self.directory_path.replace('../pdf_TextMachina', '../pdf_TextMachina/excel'), 'combined_data.xlsx')
        combined_data.to_excel(combined_output_path, index=False)

if __name__ == "__main__":
    pipeline = PDFProcessor('../pdf_TextMachina')
    pipeline.process_directory()
