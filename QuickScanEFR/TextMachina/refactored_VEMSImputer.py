
import os
import pandas as pd

class VEMSCorrector:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.data = None

    def load_workbook(self):
        self.data = pd.read_excel(self.file_path)

    @staticmethod
    def is_percent_value_usable(value):
        return isinstance(value, (int, float)) and 0 <= value <= 100

    def compute_missing_vems(self):
        vems_row = self.data[self.data[0] == 'VEMS'].iloc[0]
        for idx, val in enumerate(vems_row[1:]):
            if not self.is_percent_value_usable(val):
                self.data.iat[vems_row.name, idx + 1] = None

    def save_data(self, save_path=None):
        save_path = save_path or self.file_path.replace(".xlsx", ".xlsx")
        self.data.to_excel(save_path, index=False)

    @staticmethod
    def correct_multiple_docs(directory_path):
        for filename in os.listdir(directory_path):
            if filename.endswith("_cvf.xlsx"):
                file_path = os.path.join(directory_path, filename)
                corrector = VEMSCorrector(file_path)
                corrector.load_workbook()
                corrector.compute_missing_vems()
                corrector.save_data()
