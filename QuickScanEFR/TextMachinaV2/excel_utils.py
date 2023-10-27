import os
import pandas as pd
from openpyxl import load_workbook

# Utility functions for common tasks across modules

def load_excel_workbook(file_path):
    """Load the Excel workbook and combine sheets horizontally."""
    workbook = pd.read_excel(file_path, None)
    
    # Combine data from all sheets horizontally
    dataframes = [workbook[sheet] for sheet in workbook.keys()]
    combined_data = pd.concat(dataframes, axis=1)
    
    return combined_data

def save_excel_data(data, file_path, suffix=None):
    """Save data to an Excel file. If suffix is provided, modify the file name accordingly."""
    if suffix:
        base, ext = os.path.splitext(file_path)
        save_path = f"{base}_{suffix}{ext}"
    else:
        save_path = file_path
        
    data.to_excel(save_path, index=False)

def get_processed_file_name(file_path, suffix):
    """Generate a file name based on the provided suffix."""
    base, ext = os.path.splitext(file_path)
    return f"{base}_{suffix}{ext}"

def log_processed_file(original_file, processed_file, log_dict):
    """Log the processed file against the original in a dictionary."""
    log_dict[original_file] = processed_file

# Initialize an empty log dictionary to keep track of processed files
processed_file_log = {}
