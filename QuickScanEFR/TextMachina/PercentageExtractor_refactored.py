import os
import warnings
import pandas as pd
warnings.simplefilter(action='ignore', category=UserWarning)
import math

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
        flag = "clean"
        for value in row:
            if isinstance(value, str) and "%" in value:
                # Find the position of the first percentage sign
                percent_index = value.index('%')

                # Work backwards to determine where the numeric percentage starts
                start_index = percent_index
                    
                while start_index > 0 and not value[start_index - 1].isspace():
                    start_index -= 1
                # Split the string based on the starting index of the numeric percentage

                # Remove trailing spaces
                cleaned_value = value[:start_index].rstrip()
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
            
            # Always add the original row
            result_data.append([label] + cleaned_data)
            
            
            is_empty_row = all((not val or val == "" or (isinstance(val, float) and math.isnan(val))) for val in cleaned_data[1:])
            # If after extraction, there's percentage data but no cleaned data (excluding the label), drop the cleaned data row
            if any(percentage_data) and is_empty_row:  
                #print(cleaned_data, percentage_data)
                result_data.pop()  # Remove the last added (cleaned data) row
                
            # Add the percentage row if it contains any percentages
            if any(percentage_data):
                result_data.append([label + "%"] + percentage_data)


        result_df = pd.DataFrame(result_data, columns=df.columns)
        
        return result_df




    @classmethod
    def extract_percentages_from_excel(self, directory_path):
        """
        Extracts percentage values from all Excel files in a specified directory.

        Parameters:
        - directory_path (str): Path to the directory containing the Excel files.
        """
        for filename in filter(lambda f: f.endswith(".xlsx"), os.listdir(directory_path)):
            file_path = os.path.join(directory_path, filename)

            # Read the Excel file (assuming it has only one sheet)
            df = pd.read_excel(file_path, header=None)

            # Process the dataframe
            processed_df = self.process_dataframe_for_percentages(df)

            # Save the processed dataframe back to the Excel file
            processed_df.to_excel(file_path, index=False, header=False)
