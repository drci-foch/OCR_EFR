import os
import pandas as pd
import re 

class UnitCorrector:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.data = None

    def load_workbook(self):
        self.data = pd.read_excel(self.file_path)
        
    def correct_values(self):
        correcting_row = self.data.loc[self.data.iloc[:, 0].isin(['CV', 'CVF', 'VR Plethysmo', 'CPT Plethysmo', 'VEMS'])]
        if not correcting_row.empty:
            cvf_values = correcting_row.iloc[0, 1:].values
            corrected_values = [self._correct_unit(value) for value in cvf_values]
            self.data.iloc[correcting_row.index[0], 1:] = corrected_values

    def _correct_unit(self, value):
        if pd.notnull(value):
            str_val = str(value).strip().lower()
            xlx_pattern = re.match(r"(\d)[\s]?l[\s]?(\d{2})", str_val)
            if xlx_pattern:
                return float(xlx_pattern.group(1) + "." + xlx_pattern.group(2)) * 1000
            elif "." in str_val and len(str_val.split(".")[0]) == 1 and len(str_val.split(".")[1]) == 2:
                return float(str_val) * 1000
            else:
                numeric_part = ''.join(filter(str.isdigit, str_val))
                unit = ''.join(filter(str.isalpha, str_val))
                return float(numeric_part) * 1000 if unit in ["l", "liters"] else float(numeric_part)
        return value

    def save_corrected_data(self, save_path=None):
        save_path = save_path or self.file_path.replace(".xlsx", ".xlsx")
        self.data.to_excel(save_path, index=False)

    @staticmethod
    def correct_multiple_docs(directory_path):
        for filename in os.listdir(directory_path):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(directory_path, filename)
                corrector = UnitCorrector(file_path)
                corrector.load_workbook()
                corrector.correct_values()
                corrector.save_corrected_data()
