import os
import pandas as pd
import re 

class UnitCorrector:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.workbook = None
        self.combined_data = None

    def load_workbook(self):
        """Load the Excel workbook."""
        # Load the workbook
        self.workbook = pd.read_excel(self.file_path, None)

        # Combine data from all sheets horizontally
        dataframes = [self.workbook[sheet] for sheet in self.workbook.keys()]
        self.combined_data = pd.concat(dataframes, axis=1)
        
    def correct_values(self):
        """Correct values to ensure they are in mL."""
        # Extracting rows related to CVF from the combined data
        correcting_row = self.combined_data.loc[self.combined_data[0].isin(['CV', 'CVF', 'VR Plethysmo', 'CPT Plethysmo', 'VEMS'])]

        # Check if CVF row is not empty
        if not correcting_row.empty:
            cvf_values = correcting_row.iloc[0, 1:].values

            # Correct values to milliliters
            corrected_values = []
            for val in cvf_values:
                if pd.notnull(val):
                    str_val = str(val).strip().lower()
                    
                    # Check for format like "2L50"
                    xlx_pattern = re.match(r"(\d)[\s]?l[\s]?(\d{2})", str_val)
                    if xlx_pattern:
                        corrected_value = float(xlx_pattern.group(1) + "." + xlx_pattern.group(2)) * 1000
                    # Check for format like "2.50"
                    elif "." in str_val and len(str_val.split(".")[0]) == 1 and len(str_val.split(".")[1]) == 2:
                        corrected_value = float(str_val) * 1000
                    else:
                        # Extract the numeric part from the string
                        numeric_part = ''.join(filter(str.isdigit, str_val))

                        # Check for unit in the value (L or liters)
                        unit = ''.join(filter(str.isalpha, str_val))

                        # Convert to mL by multiplying by 1000 if the unit is "l" or "liters"
                        if unit in ["l", "liters"]:
                            corrected_value = float(numeric_part) * 1000
                        else:
                            # Assume it's in mL if no explicit unit found
                            corrected_value = float(numeric_part)

                    corrected_values.append(corrected_value)
                else:
                    corrected_values.append(val)

            # Update the dataframe with the corrected CVF values
            self.combined_data.iloc[correcting_row.index[0], 1:] = corrected_values




    def save_corrected_data(self, save_path=None):
        """Save the corrected data back to the Excel file."""
        if not save_path:
            save_path = self.file_path.replace(".xlsx", "_cvf.xlsx")
        self.combined_data.to_excel(save_path, index=False)

    @staticmethod
    def correct_multiple_docs(directory_path):
        """Correct CVF values for multiple Excel documents in a specified directory."""
        for filename in os.listdir(directory_path):
            if filename.endswith("_corrected.xlsx"):
                file_path = os.path.join(directory_path, filename)
                corrector = UnitCorrector(file_path)
                corrector.load_workbook()
                corrector.correct_values()
                corrector.save_corrected_data()
