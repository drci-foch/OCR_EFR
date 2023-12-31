import os
import re
import pandas as pd

class ExcelCleanup:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.combined_data = None

    def load_workbook(self):
        """Load the Excel workbook."""
        self.combined_data = pd.read_excel(self.file_path)

    @staticmethod
    def format_percentage(value, is_first_column=False):
        """Format the value to the desired percentage format."""
        # If it's the first column, retain the '%' symbol
        if is_first_column:
            return value
        
        # Convert string percentages like "90%" to 90 only if they strictly match the "X%" format
        if isinstance(value, str) and re.match(r"^\d+(\.)?%$", value):  # This regex also allows for decimals like "64.56%"
            return float(value.replace('%', ''))
        
        # Convert float percentages like 0.9 to 90
        elif isinstance(value, (float, int)) and 0 < value <= 1:
            return value * 100
        
        else: 
            return value
        
        

    def clean_percentage_values(self):
        """Clean percentage values for rows with first column ending in '%'."""
    
        # Filter rows where the first column value ends with '%'
        rows_to_process = self.combined_data[self.combined_data.iloc[:, 0].str.endswith('%', na=False)]
        
        # Process every column for the filtered rows
        for _, row in rows_to_process.iterrows():
            for idx, col in enumerate(self.combined_data.columns):
                self.combined_data.at[row.name, col] = ExcelCleanup.format_percentage(self.combined_data.at[row.name, col], is_first_column=(idx == 0))

    @staticmethod
    def is_convertible_to_float(value):
        """Check if a string value can be converted to a float."""
        try:
            float(value)
            return True
        except ValueError:
            return False

    def save_cleaned_data(self, save_path=None):
        """Save the cleaned data back to the Excel file."""
        if not save_path:
            save_path = self.file_path.replace(".xlsx", "_final.xlsx")
        self.combined_data.to_excel(save_path, index=False)

    @staticmethod
    def clean_multiple_docs(directory_path):
        """Clean percentage values for multiple Excel documents in a specified directory."""
        for filename in os.listdir(directory_path):
            if filename.endswith("_cvf.xlsx"):
                file_path = os.path.join(directory_path, filename)
                cleaner = ExcelCleanup(file_path)
                cleaner.load_workbook()
                cleaner.clean_percentage_values()
                cleaner.save_cleaned_data()
