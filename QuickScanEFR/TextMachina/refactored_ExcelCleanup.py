
import os
import pandas as pd
import re 

class ExcelCleanup:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.data = None

    def load_workbook(self):
        self.data = pd.read_excel(self.file_path, header=None)

    def transpose_data(self):
        # Reset the index to make sure the dates become a regular column
        df_transposed = self.data.transpose()
        df_transposed.columns = df_transposed.iloc[0]
        df_transposed = df_transposed.drop(df_transposed.index[0])
        # Rename the first column to "date"
        df_transposed.columns.values[0] = 'date'
       
        self.data = df_transposed
        return self.data

    @staticmethod
    def format_percentage(value, is_first_column=False):
        if is_first_column or not isinstance(value, str):
            return value
        if re.match(r"^\d+(\.\d+)?%$", value):
            return float(value.replace('%', ''))
        elif isinstance(value, (float, int)) and 0 < value <= 1:
            return value * 100
        return value

    def clean_percentage_values(self):
        rows_to_process = self.data[self.data.iloc[:, 0].str.endswith('%', na=False)]
        for _, row in rows_to_process.iterrows():
            for idx, col in enumerate(self.data.columns):
                self.data.at[row.name, col] = ExcelCleanup.format_percentage(
                    self.data.at[row.name, col], 
                    is_first_column=(idx == 0)
                )

    def save_cleaned_data(self, output_path):
        """Save the cleaned data to the specified output path."""
        output_file = os.path.join(output_path, os.path.basename(self.file_path))
        self.data.to_excel(output_file, index=False)
        print(f"Saved cleaned data to: {output_file}")

    @staticmethod
    def clean_multiple_docs(directory_path, output_directory):
        """Process multiple documents with specified input and output directories."""
        # Ensure output directory exists
        os.makedirs(output_directory, exist_ok=True)
        
        for filename in os.listdir(directory_path):
            if filename.endswith(".xlsx"):
                print(f"--------------------------Processing {filename}--------------------------")
                try:
                    file_path = os.path.join(directory_path, filename)
                    
                    # Create and process cleaner
                    cleaner = ExcelCleanup(file_path)
                    cleaner.load_workbook()
                    cleaner.clean_percentage_values()
                    cleaner.transpose_data()
                    cleaner.save_cleaned_data(output_directory)
                    
                    print(f"Processed {filename} successfully.")
                except Exception as e:
                    print(f"Error processing {filename}: {e}")