
import os
import pandas as pd

class VEMSCorrector:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.combined_data = None

    def load_workbook(self):
        self.combined_data = pd.read_excel(self.file_path)

    @staticmethod
    def is_percent_value_usable(value):
        return isinstance(value, (int, float)) and 0 <= value <= 100

    def compute_missing_vems(self):
        vems_row = self.combined_data[self.combined_data[0] == 'VEMS'].iloc[0]
        for idx, val in enumerate(vems_row[1:]):
            if not self.is_percent_value_usable(val):
                self.combined_data.iat[vems_row.name, idx + 1] = None

    def save_combined_data(self, save_path=None):
        save_path = save_path or self.file_path.replace(".xlsx", "_vems.xlsx")
        self.combined_data.to_excel(save_path, index=False)

    @staticmethod
    def correct_multiple_docs(directory_path):
        for filename in os.listdir(directory_path):
            if filename.endswith("_cvf.xlsx"):
                file_path = os.path.join(directory_path, filename)
                corrector = VEMSCorrector(file_path)
                corrector.load_workbook()
                corrector.compute_missing_vems()
                corrector.save_combined_data()
