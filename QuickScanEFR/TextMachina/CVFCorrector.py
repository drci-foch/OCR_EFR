import os
import pandas as pd

class CVFCorrector:
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
        
    def correct_cvf_values(self):
        """Correct CVF values to ensure they are in mL."""
        # Extracting rows related to CVF from the combined data
        cvf_row = self.combined_data[self.combined_data[0] == 'CVF']
        
        # Check if CVF row is not empty
        if not cvf_row.empty:
            cvf_values = cvf_row.iloc[0, 1:].values
        
            # Correct values that are likely in L
            corrected_values = []
            for val in cvf_values:
                if pd.notnull(val) and len(str(val).split('.')[0]) == 1:
                    corrected_values.append(val * 1000)
                else:
                    corrected_values.append(val)
        
            # Update the dataframe with the corrected CVF values
            self.combined_data.iloc[cvf_row.index[0], 1:] = corrected_values

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
                corrector = CVFCorrector(file_path)
                corrector.load_workbook()
                corrector.correct_cvf_values()
                corrector.save_corrected_data()
