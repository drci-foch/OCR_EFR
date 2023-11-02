
import os
import pandas as pd
import re 

class ExcelCleanup:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.data = None

    def load_workbook(self):
        self.data = pd.read_excel(self.file_path)

    def transpose_data(self):
        # Reset the index to make sure the dates become a regular column

        df_transposed = self.data.transpose()
        df_transposed.columns = df_transposed.iloc[0]
        df_transposed = df_transposed.drop(df_transposed.index[0])
        
        self.data = df_transposed
        return self.data  # Return for inspection


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
                self.data.at[row.name, col] = ExcelCleanup.format_percentage(self.data.at[row.name, col], is_first_column=(idx == 0))

    def save_cleaned_data(self, save_path=None):
        save_path = save_path or self.file_path.replace(".xlsx", ".xlsx")
        self.data.to_excel(save_path, index=True)

    @staticmethod
    def clean_multiple_docs(directory_path):
        for filename in os.listdir(directory_path):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(directory_path, filename)
                cleaner = ExcelCleanup(file_path)
                cleaner.load_workbook()
                cleaner.clean_percentage_values()
                cleaner.transpose_data()  
                cleaner.save_cleaned_data()
