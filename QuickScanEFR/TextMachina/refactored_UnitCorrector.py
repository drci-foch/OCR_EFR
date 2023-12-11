import os
import pandas as pd
import re 

class UnitCorrector:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.data = None

    def load_workbook(self):
        self.data = pd.read_excel(self.file_path, header=None)
        
    def correct_values(self):
        for index, row in self.data.iterrows():
            if row[0] in ['CV', 'CVF', 'VR Plethysmo', 'CPT Plethysmo', 'VEMS', 'AA/O²']:
                corrected_values = [self._correct_unit(value) for value in row[1:]]
                self.data.iloc[index, 1:] = corrected_values
                print(self.data)


    def _correct_unit(self, value):
        if pd.notnull(value):
            str_val = str(value).strip().lower()
            # Remove occurrences of 'ml', 'mL', or 'ML'
            str_val = re.sub(r'\bml\b', '', str_val, flags=re.IGNORECASE)
            
            xlx_pattern = re.match(r"(\d{1,2})[\s]?[lL][\s]?(\d{1,2})", str_val)
            
            if xlx_pattern:
                print(xlx_pattern.group(1), xlx_pattern.group(2))
                return float(xlx_pattern.group(1) + "." + xlx_pattern.group(2)) * 1000
            elif "." in str_val and len(str_val.split(".")[0]) == 1 and len(str_val.split(".")[1]) == 2:
                return float(str_val) * 1000
            elif "," in str_val and len(str_val.split(",")[0]) == 1 and len(str_val.split(",")[1]) == 2:
                str_val = str_val.replace(',', '.')
                return float(str_val) * 1000
            elif " " in str_val and len(str_val.split(" ")[0]) == 1 and len(str_val.split(" ")[1]) == 2:
                return float(str_val) * 1000
            else:
                numeric_part = ''.join(filter(str.isdigit, str_val))
                unit = ''.join(filter(str.isalpha, str_val))
                return float(numeric_part) * 1000 if unit in ["l", "litre","litres", "L"] else float(numeric_part)
        return value

    def save_corrected_data(self, save_path=None):
        """
        Save the corrected data to an Excel file in a specific folder.

        Parameters:
        - save_path (str, optional): Path for saving the corrected data.
        """
        # Define the new folder name
        new_folder = 'excel_correct_unit'

        # Extract the directory from the existing file path
        dir_path = os.path.dirname(self.file_path)

        # Create the new directory path
        new_dir_path = os.path.join(dir_path, new_folder)

        # Create the new folder if it doesn't exist
        if not os.path.exists(new_dir_path):
            os.makedirs(new_dir_path)

        # Construct the new file path
        if save_path is None:
            new_file_name = os.path.basename(self.file_path)
            save_path = os.path.join(new_dir_path, new_file_name)

        # Save the file
        self.data.to_excel(save_path, index=False, header=None)

    @staticmethod
    def correct_multiple_docs(directory_path):
        for filename in os.listdir(directory_path):
            print(f"---------------------Processing {filename}---------------------")
            if filename.endswith(".xlsx"):
                try:
                    file_path = os.path.join(directory_path, filename)
                    corrector = UnitCorrector(file_path)
                    corrector.load_workbook()
                    corrector.correct_values()
                    corrector.save_corrected_data()
                    print(f"Processed and saved: {filename}")
                except Exception as e:
                    print(f"Error processing {filename}: {e}")