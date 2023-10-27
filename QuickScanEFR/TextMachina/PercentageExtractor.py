import os
import warnings 
import pandas as pd
warnings.simplefilter(action='ignore', category=UserWarning)

class PercentageExtractor:
    """
    A class responsible for extracting percentage values from Excel rows and processing dataframes for percentages.
    """
    @staticmethod
    def extract_percentages_from_row(row):
        """
        Extracts percentage values from a row of data.

        Parameters:
        - row (Series): A row of data containing values that might be in the format "value percentage".

        Returns:
        - tuple: Two lists, one with cleaned values and another with extracted percentages.
        """
        cleaned_values = []
        percentage_values = []

        for value in row:
            if isinstance(value, str) and "%" in value:
                # Find the position of the first percentage sign
                percent_index = value.index('%')
                
                # Work backwards to determine where the numeric percentage starts
                start_index = percent_index
                while start_index > 0 and not value[start_index - 1].isspace():
                    start_index -= 1
                
                # Split the string based on the starting index of the numeric percentage
                cleaned_value = value[:start_index].rstrip()  # Remove trailing spaces
                percent_value = value[start_index:].strip()
                
                cleaned_values.append(cleaned_value)
                percentage_values.append(percent_value)
            else:
                cleaned_values.append(value)
                percentage_values.append(None)

        return cleaned_values, percentage_values



    @staticmethod
    def process_dataframe_for_percentages(df):
        """
        Processes a dataframe and extracts percentage rows from it.

        Parameters:
        - df (DataFrame): The input dataframe to process.

        Returns:
        - DataFrame: A dataframe with original and extracted percentage rows.
        """
        result_data = []

        for _, row in df.iterrows():
            label = str(row[0])
            data = row[1:]
            
            cleaned_data, percentage_data = PercentageExtractor.extract_percentages_from_row(data)
            
            # Add the original row
            result_data.append([label] + cleaned_data)
            
            # Add the percentage row immediately after the original row, only if it contains percentages
            if any(percentage_data):
                result_data.append([label + "%"] + percentage_data)

        result_df = pd.DataFrame(result_data, columns=df.columns)
        
        return result_df

    @classmethod
    def extract_percentages_from_excel(cls, directory):
        """
        Extracts percentage values from all Excel files in a specified directory that have been processed by DateFormatter.

        Parameters:
        - directory (str): Path to the directory containing the Excel files.
        """
        for filename in os.listdir(directory):
            # Look for files that have been processed by DateFormatter
            if filename.endswith("_cleaned.xlsx"):
                file_path = os.path.join(directory, filename)
                
                with pd.ExcelFile(file_path) as xls:
                    all_sheets = {}
                    
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name)
                        processed_df = cls.process_dataframe_for_percentages(df)
                        all_sheets[sheet_name] = processed_df
                        
                output_file_path = os.path.join(directory, filename)
                with pd.ExcelWriter(output_file_path) as writer:
                    for sheet_name, data in all_sheets.items():
                        data.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"Processed: {filename}")

