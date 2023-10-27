import os
from excel_utils import load_excel_workbook, save_excel_data, get_processed_file_name

class CVFCorrectorRefactored:
    def __init__(self, file_path=None):
        self.file_path = file_path

    def correct_cvf_values(self, data):
        """Correct CVF values in the combined data."""
        for col in data.columns:
            data[col] = data[col].apply(lambda x: x * 1000 if isinstance(x, (int, float)) and len(str(int(x))) == 1 else x)
        return data

    def correct_multiple_docs(self, directory_path):
        """Correct CVF values for multiple Excel documents in a specified directory."""
        for filename in os.listdir(directory_path):
            if filename.endswith("_modified_corrected_vems.xlsx"):
                file_path = os.path.join(directory_path, filename)
                data = load_excel_workbook(file_path)
                corrected_data = self.correct_cvf_values(data)
                save_path = get_processed_file_name(file_path, "cvf")
                save_excel_data(corrected_data, save_path)
                print(f"Processed CVF for {file_path} and saved as {save_path}")
