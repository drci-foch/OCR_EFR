import pandas as pd
import os

def concatenate_excel_files(directory_path):
    # List all files in the given directory
    all_files = os.listdir(directory_path)
    
    # Filter out the Excel files
    excel_files = [f for f in all_files if f.endswith('.xlsx')]
    
    # Group files by ID
    files_by_id = {}
    for file in excel_files:
        file_id = file.split('_')[0]
        if file_id not in files_by_id:
            files_by_id[file_id] = []
        files_by_id[file_id].append(file)
    
    # Concatenate files with the same ID
    for file_id, files in files_by_id.items():
        dfs = []
        for file in files:
            dfs.append(pd.read_excel(os.path.join(directory_path, file)))
        
        concatenated_df = pd.concat(dfs, axis=0, ignore_index=True)
        
        # Save concatenated file
        concatenated_df.to_excel(os.path.join(directory_path, f"{file_id}.xlsx"), index=False)

