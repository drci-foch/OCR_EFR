
import os
import pandas as pd
import re 

class ExcelCleanup:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.combined_data = None

    def load_workbook(self):
        self.combined_data = pd.read_excel(self.file_path)

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
        rows_to_process = self.combined_data[self.combined_data.iloc[:, 0].str.endswith('%', na=False)]
        for _, row in rows_to_process.iterrows():
            for idx, col in enumerate(self.combined_data.columns):
                self.combined_data.at[row.name, col] = ExcelCleanup.format_percentage(self.combined_data.at[row.name, col], is_first_column=(idx == 0))

    def save_cleaned_data(self, save_path=None):
        save_path = save_path or self.file_path.replace(".xlsx", "_final.xlsx")
        self.combined_data.to_excel(save_path, index=False)

    @staticmethod
    def clean_multiple_docs(directory_path):
        for filename in os.listdir(directory_path):
            if filename.endswith("_cvf.xlsx"):
                file_path = os.path.join(directory_path, filename)
                cleaner = ExcelCleanup(file_path)
                cleaner.load_workbook()
                cleaner.clean_percentage_values()
                cleaner.save_cleaned_data()
