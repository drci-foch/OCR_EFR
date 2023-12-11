import pandas as pd
import os
import glob
import re

# Define your folder path
folder_path = 'C:/Users/benysar/Desktop/Github/OCR_EFR/QuickScanEFR/pdf_TextMachina/excel_cleanup'
new_folder_path = 'C:/Users/benysar/Desktop/Github/OCR_EFR/QuickScanEFR/pdf_TextMachina/excel_cleanup/concat'


# Create new folder if it doesn't exist
if not os.path.exists(new_folder_path):
    os.makedirs(new_folder_path)

# Function to extract patient ID from filename
def get_patient_id(filename):
    return filename.split('_')[0]

# Find all Excel files
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

# Group files by patient ID
patient_files = {}
for file in excel_files:
    patient_id = get_patient_id(os.path.basename(file))
    patient_files.setdefault(patient_id, []).append(file)

# Merge files for each patient
for patient_id, files in patient_files.items():
    merged_df = None
    for file in files:
        df = pd.read_excel(file)

        # Convert all columns to string to avoid data type mismatch
        df = df.astype(str)

        if merged_df is None:
            merged_df = df
        else:
            # Merge while avoiding duplicate columns
            merged_df = pd.merge(merged_df, df, how='outer')

    # Save merged DataFrame to a new Excel file
    merged_filename = os.path.join(new_folder_path, f'merged_{patient_id}.xlsx')
    merged_df.to_excel(merged_filename, index=False)

print("Merging completed!")