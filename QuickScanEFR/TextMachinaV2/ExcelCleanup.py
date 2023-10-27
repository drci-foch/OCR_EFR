import os
from excel_utils import load_excel_workbook, save_excel_data, get_processed_file_name

class ExcelCleanupRefactored:
    def __init__(self, file_path=None):
        self.file_path = file_path

    @staticmethod
    def is_convertible_to_float(value):
        """Check if a value can be converted to float."""
        try:
            float(value)
            return True
        except ValueError:
            return False

    def format_percentage(self, value):
        """Format a value to the desired percentage format."""
        if isinstance(value, str) and '%' in value:
            value = value.replace('%', '')
            if self.is_convertible_to_float(value):
                return float(value) * 100 if float(value) <= 1 else float(value)
        elif self.is_convertible_to_float(value):
            return float(value) * 100 if float(value) <= 1 else float(value)
        return value

    def clean_percentage_values(self, data):
        """Clean percentage values for the combined data."""
        for col in data.columns:
            if data[col].dtype == 'object' and any(data[col].str.endswith('%', na=False)):
                data[col] = data[col].apply(self.format_percentage)
        return data

    def clean_multiple_docs(self, directory_path):
        """Clean percentage values for multiple Excel documents in a specified directory."""
        for filename in os.listdir(directory_path):
            if filename.endswith("_cvf.xlsx"):
                file_path = os.path.join(directory_path, filename)
                data = load_excel_workbook(file_path)
                cleaned_data = self.clean_percentage_values(data)
                save_path = get_processed_file_name(file_path, "cleaned")
                save_excel_data(cleaned_data, save_path)
                print(f"Processed {file_path} and saved as {save_path}")
