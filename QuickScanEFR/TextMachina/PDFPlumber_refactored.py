import os
import shutil
os.environ['PATH'] += os.pathsep + 'C:\\Users\\benysar\\AppData\\Local\\gs10.02.1\\bin'
import camelot
import pandas as pd
from openpyxl import Workbook, load_workbook

class PDFProcessor:
    """
    A class to process PDFs in a given directory, extract tables, 
    and save them as Excel files.
    
    Attributes:
        directory_path (str): Path to the directory containing PDF files.
    """

    def __init__(self, directory_path, output_path):
        """
        Initializes the PDFProcessor with the given directory path.
        
        Args:
            directory_path (str): Path to the directory containing PDF files.
        """
        self.directory_path = directory_path
        self.output_path = output_path

    def process_pdf(self, pdf_path, output_path):
        """
        Processes a single PDF, extracts tables from it, 
        and saves them in an Excel file.
        
        Args:
            pdf_path (str): Path to the PDF file.
            output_path (str): Path to save the resulting Excel file.
        """
        try: 
            tables = camelot.read_pdf(pdf_path, pages='1-end', flavor="lattice", strip_text='\n')
            print(f"Here are the number of tables extracted for {os.path.basename(pdf_path)}:", tables)
            with pd.ExcelWriter(output_path) as writer:
                for i, table in enumerate(tables):
                    table.df.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)
        except ZeroDivisionError:
            problem_folder = os.path.join(self.directory_path, 'problem')
            os.makedirs(problem_folder, exist_ok=True)
            problem_file_path = os.path.join(problem_folder, os.path.basename(pdf_path))
            os.rename(pdf_path, problem_file_path)
            print(f"Problem with file {pdf_path}. Moved to 'problem' folder.")

    def combine_data_horizontally(self, output_path):
        """
        Combines data from multiple sheets in an Excel file horizontally.

        Args:
            output_path (str): Path to the Excel file to save the combined data.

        Returns:
            pd.DataFrame: Combined and pivoted DataFrame.
        """
        xls = pd.ExcelFile(output_path)
        dfs, original_order_date, original_order_measure = self.read_excel_sheets(xls)
        combined_df = self.concat_dataframes(dfs)
        deduplicated_df = self.drop_duplicates(combined_df)
        pivoted_df = self.pivot_dataframe(deduplicated_df, original_order_date, original_order_measure)
        print("Combination of data successful.")
        return pivoted_df

    def read_excel_sheets(self, xls):
        """
        Reads and processes sheets from an Excel file.

        Args:
            xls (pd.ExcelFile): Excel file object.

        Returns:
            list: List of DataFrames, original_order_date, original_order_measure.
        """
        dfs = []
        original_order_date = []
        original_order_measure = []

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            # Check if there's only one column in the sheet
            if len(df.columns) == 1:
                continue  # Skip processing this sheet
            df, date, measure = self.process_dataframe(df)
            original_order_date += date
            original_order_measure += measure
           
            dfs.append(df)

        return dfs, original_order_date, original_order_measure

    def process_dataframe(self, df):
        """
        Process a DataFrame from an Excel sheet.

        Args:
            df (pd.DataFrame): DataFrame to process.

        Returns:
            pd.DataFrame: Processed DataFrame.
            list: Original order of dates.
            list: Original order of measures.
        """
        # Set the first row as column headers (dates)
        df.columns = df.iloc[1]
        # Remove the first row from the DataFrame
        df = df.drop(df.index[0])
        original_order_date = list(df.iloc[0].dropna())
        original_order_measure = list(df.iloc[:, 0])
        
        # Rename the first column to 'Measure' and set it as the index
        df = df.rename(columns={df.columns[0]: "Measure"})
        df = df.set_index("Measure")

        # Melt the DataFrame to long format
        df_long = df.reset_index().melt(id_vars="Measure", var_name='Date', value_name='Value')

        # Extract the first element from the tuple in the "Measure" column
        df_long['Measure'] = df_long['Measure'].apply(lambda x: x[0] if isinstance(x, tuple) else x)

        return df_long, original_order_date, original_order_measure

    def concat_dataframes(self, dfs):
        """
        Concatenates a list of DataFrames.

        Args:
            dfs (list): List of DataFrames.

        Returns:
            pd.DataFrame: Combined DataFrame.
        """
        combined_df = pd.concat(dfs, ignore_index=True)
        return combined_df

    def drop_duplicates(self, combined_df):
        """
        Drops duplicate rows in a DataFrame.

        Args:
            combined_df (pd.DataFrame): DataFrame to deduplicate.

        Returns:
            pd.DataFrame: Deduplicated DataFrame.
        """
        deduplicated_df = combined_df.drop_duplicates(subset=['Measure', 'Date'])
        deduplicated_df = deduplicated_df.dropna(subset=['Measure', 'Date'])
        return deduplicated_df

    def pivot_dataframe(self, deduplicated_df, original_order_date, original_order_measure):
        """
        Pivots a DataFrame.

        Args:
            deduplicated_df (pd.DataFrame): Deduplicated DataFrame.
            original_order_date (list): Original order of dates.
            original_order_measure (list): Original order of measures.

        Returns:
            pd.DataFrame: Pivoted DataFrame.
        """
        pivoted_df = deduplicated_df.pivot(index='Measure', columns='Date', values='Value')
        original_order_measure = [item for item in list(dict.fromkeys(original_order_measure)) if not pd.isna(item)]
        for measure in original_order_measure:
            if measure not in pivoted_df.index:
                pivoted_df.loc[measure] = None  # Handling missing measure gracefully
        # Check if dates are in the columns before reordering
        existing_dates = [date for date in original_order_date if date in pivoted_df.columns]
        pivoted_df = pivoted_df[existing_dates].reset_index()
        pivoted_df = pivoted_df.set_index('Measure')
        pivoted_df = pivoted_df.loc[original_order_measure]
        pivoted_df.index.name = None
        return pivoted_df

    def process_directory(self):
        for filename in os.listdir(self.directory_path):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(self.directory_path, filename)
                output_excel_path = os.path.join(self.output_path, filename.replace('.pdf', '.xlsx'))
                self.process_and_combine(pdf_path, output_excel_path)


    def process_and_combine(self, pdf_path, output_path):
        """
        Processes a PDF and combines the extracted tables horizontally.

        Args:
            pdf_path (str): Path to the PDF file.
            output_path (str): Path to save the resulting Excel file.
        """
        try:
            #self.process_pdf(pdf_path, output_path)
            combined_data = self.combine_data_horizontally(output_path)
            combined_data.to_excel(output_path)  # Provide the 'output_path' here
            pass
        except ZeroDivisionError:
            problem_folder = os.path.join(self.directory_path, 'problem')
            os.makedirs(problem_folder, exist_ok=True)
            problem_file_path = os.path.join(problem_folder, os.path.basename(pdf_path))
            os.rename(pdf_path, problem_file_path)
            print(f"Problem with file {pdf_path}. Moved to 'problem' folder.")
        except Exception as e:
            problem_folder = os.path.join(self.directory_path, 'exception')
            os.makedirs(problem_folder, exist_ok=True)
            
            # Move the problematic file
            problem_file_path = os.path.join(problem_folder, os.path.basename(pdf_path))
            os.rename(pdf_path, problem_file_path)

            # Log the exception in an Excel file
            excel_file = os.path.join(problem_folder, 'exceptions_log.xlsx')
            if not os.path.exists(excel_file):
                wb = Workbook()
                ws = wb.active
                ws.append(["Folder Name", "Exception"])
            else:
                wb = load_workbook(excel_file)
                ws = wb.active

            ws.append([os.path.basename(pdf_path), str(e)])
            wb.save(excel_file)
            print(f"An error occurred while processing {pdf_path}: {e}. Moved to 'exception' folder.")


# if __name__ == "__main__":
#     pipeline = PDFProcessor(directory_path='../pdf_TextMachina/excel/0189149780_2015-12-31_D6DB5B95-5605-46D3-B653-08103D0A6EC4.pdf',output_path= '../pdf_TextMachina/excel/outputs.xlsx')
#     pipeline.process_pdf()