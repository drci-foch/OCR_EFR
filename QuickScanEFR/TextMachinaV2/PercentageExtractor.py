import os
import re
import pandas as pd
from excel_utils import load_excel_workbook, get_processed_file_name, save_excel_data

# Refactoring the PercentageExtractor.py module

class PercentageExtractorRefactored:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.data = None

    def extract_percentages_from_row(self, row):
        """Extract percentage values from a row."""
        new_row = row.copy()
        for idx, val in enumerate(row):
            if isinstance(val, str) and "%" in val:
                # Split the value into two parts: absolute value and percentage value
                parts = re.split(r'\s*(?=%)', val, maxsplit=1)
                new_row[idx] = parts[0]
                if len(parts) > 1:
                    new_row.at["Percentage"] = parts[1]
        return new_row

    def process_dataframe_for_percentages(self, df):
        """Process dataframe to extract percentage rows and place them immediately after the original rows."""
        new_rows = []
        for idx, row in df.iterrows():
            new_row = self.extract_percentages_from_row(row)
            new_rows.append(new_row)
        return pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

    def extract_percentages_from_excel(self, directory_path):
        """Extract percentage values from all Excel files in a specified directory."""
        for filename in os.listdir(directory_path):
            if filename.endswith("_cleaned.xlsx"):
                file_path = os.path.join(directory_path, filename)
                self.data = load_excel_workbook(file_path)
                processed_data = self.process_dataframe_for_percentages(
                    self.data)
                save_path = get_processed_file_name(file_path, "modified")
                save_excel_data(processed_data, save_path)
                print(f"Processed {file_path} and saved as {save_path}")

