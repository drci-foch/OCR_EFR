import pandas as pd
import os

def extract_ipp_from_filename(filename):
    """Extract IPP from filename by removing the .xlsx extension"""
    return filename.replace('.xlsx', '')

def process_excel_files(directory):
    # List to store all dataframes
    all_dfs = []
    
    # Process each Excel file in the directory
    for filename in os.listdir(directory):
        if not filename.endswith('.xlsx'):
            continue
            
        # Get IPP from filename
        ipp = extract_ipp_from_filename(filename)
        
        try:
            # Read the Excel file
            file_path = os.path.join(directory, filename)
            df = pd.read_excel(file_path)
            
            # Add IPP column at the beginning
            df.insert(0, 'IPP', ipp)
            
            # Append to list of dataframes
            all_dfs.append(df)
            print(f"Processed: {filename}")
            
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")
    
    if not all_dfs:
        print("No Excel files were processed")
        return None
    
    # Concatenate all dataframes
    combined_df = pd.concat(all_dfs, ignore_index=True)
    return combined_df

def main():
    # Directory containing the renamed Excel files
    directory = './renamed_files'  # Update this path if your files are elsewhere
    
    # Process files
    print("Starting file processing...")
    combined_df = process_excel_files(directory)
    
    if combined_df is not None:
        # Save the combined data
        output_filename = 'combined_data.xlsx'
        combined_df.to_excel(output_filename, index=False)
        
        # Print summary
        print(f"\nProcessing complete:")
        print(f"Total rows in combined file: {len(combined_df)}")
        print(f"Total unique IPPs: {combined_df['IPP'].nunique()}")
        print(f"Columns in final file: {', '.join(combined_df.columns)}")
        print(f"\nCombined data saved to: {output_filename}")

if __name__ == "__main__":
    main()
