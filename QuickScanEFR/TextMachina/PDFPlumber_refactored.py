import os
import shutil
os.environ['PATH'] += os.pathsep + r'C:\Program Files\gs\gs10.04.0\bin'
import camelot
import pandas as pd
from openpyxl import Workbook, load_workbook

class PDFProcessor:
    def __init__(self, directory_path, output_path):
        self.directory_path = os.path.abspath(directory_path)
        self.output_path = os.path.abspath(output_path)
        print(f"Initialized with input directory: {self.directory_path}")
        print(f"Output directory: {self.output_path}")
        os.makedirs(self.output_path, exist_ok=True)

    def process_pdf(self, pdf_path, output_path):
        try:
            if not os.path.exists(pdf_path):
                raise FileNotFoundError(f"PDF file not found: {pdf_path}")
                
            print(f"Processing PDF: {pdf_path}")
            tables = camelot.read_pdf(pdf_path, pages='1-end', flavor="lattice", strip_text='\n')
            print(f"Extracted {len(tables)} tables from {os.path.basename(pdf_path)}")
            
            with pd.ExcelWriter(output_path) as writer:
                for i, table in enumerate(tables):
                    table.df.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)
            print(f"Saved tables to: {output_path}")
            return True
            
        except Exception as e:
            print(f"Error processing PDF: {str(e)}")
            self.handle_error(pdf_path, 'exception', str(e))
            return False

    def process_and_combine(self, pdf_path, output_path):
        """
        Processes a PDF and combines the extracted tables horizontally.

        Args:
            pdf_path (str): Path to the PDF file.
            output_path (str): Path to save the resulting Excel file.
        """
        try:
            print(f"Starting processing of {pdf_path}")
            # Uncommented the PDF processing step
            self.process_pdf(pdf_path, output_path)
            print(f"PDF processed, combining data from {output_path}")
            combined_data = self.combine_data_horizontally(output_path)
            print("Data combined, saving final result")
            combined_data.to_excel(output_path)
            print(f"Processing complete for {pdf_path}")
            return True
        except Exception as e:
            print(f"Error processing {pdf_path}: {str(e)}")
            self.handle_error(pdf_path, 'exception', str(e))
            return False

    def process_directory(self):
        print("\nStarting directory processing...")
        print(f"Looking for PDFs in: {self.directory_path}")
        
        if not os.path.exists(self.directory_path):
            print(f"Input directory does not exist: {self.directory_path}")
            return []
            
        processed_files = []
        pdf_files = [f for f in os.listdir(self.directory_path) if f.endswith('.pdf')]
        print(f"Found {len(pdf_files)} PDF files")
        
        for filename in pdf_files:
            pdf_path = os.path.join(self.directory_path, filename)
            output_excel_path = os.path.join(self.output_path, filename.replace('.pdf', '.xlsx'))
            
            print(f"\nProcessing file: {filename}")
            if self.process_and_combine(pdf_path, output_excel_path):
                processed_files.append(filename)
                print(f"Successfully processed {filename}")
            else:
                print(f"Failed to process {filename}")
        
        print(f"\nProcessed {len(processed_files)} files successfully")
        return processed_files

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

# if __name__ == "__main__":
#     pipeline = PDFProcessor(directory_path='../pdf_TextMachina/excel/0189149780_2015-12-31_D6DB5B95-5605-46D3-B653-08103D0A6EC4.pdf',output_path= '../pdf_TextMachina/excel/outputs.xlsx')
#     pipeline.process_pdf()