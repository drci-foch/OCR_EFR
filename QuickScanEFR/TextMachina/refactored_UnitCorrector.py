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

    def save_corrected_data(self, output_path):
        """Save the corrected data to an Excel file."""
        if not os.path.exists(output_path):
            os.makedirs(output_path)
            
        output_file = os.path.join(output_path, os.path.basename(self.file_path))
        self.data.to_excel(output_file, index=False, header=None)
        print(f"Saved corrected data to: {output_file}")

    @staticmethod
    def correct_multiple_docs(directory_path, output_path):
        """Process multiple documents and save to specified output path."""
        # Ensure output directory exists
        os.makedirs(output_path, exist_ok=True)
        
        for filename in os.listdir(directory_path):
            print(f"---------------------Processing {filename}---------------------")
            if filename.endswith(".xlsx"):
                try:
                    # Full paths for input and output
                    input_file = os.path.join(directory_path, filename)
                    
                    # Create and process corrector
                    corrector = UnitCorrector(input_file)
                    corrector.load_workbook()
                    corrector.correct_values()
                    corrector.save_corrected_data(output_path)
                    print(f"Successfully processed {filename}")
                except Exception as e:
                    print(f"Error processing {filename}: {e}")