import os
import shutil
import pandas as pd
import math
import re


class PercentageExtractor:
    """
    A class responsible for extracting percentage values from Excel rows and processing dataframes for percentages.
    """
    @staticmethod
    def extract_percentages_from_row(row):
        cleaned_values = []
        percentage_values = []

        for value in row:
            str_value = str(value)
            # First, check for existing percentage patterns
            matches = re.findall(r'(\d+(\/\d+)?\s*[%ù])', str_value)

            # Check for a two-number pattern where the second number is a two-digit number
            two_number_match = re.search(r'(\d+)\s+(\d{2,}(?:[.,]\d+)?%?)', str_value)
            if two_number_match and not matches:
                
                number1, number2 = two_number_match.groups()
                print(number2)     
                cleaned_values.append(str_value.replace(f"{number1} {number2}", number1).strip())
                percentage_values.append(f"{number2}%")
                
            elif matches:
                # Handle matched percentage expressions
                percent_value = ' '.join([match[0] for match in matches]).replace('ù', '%')
                
                cleaned_value = str_value
                for match in matches:
                    cleaned_value = cleaned_value.replace(match[0], '').strip()
                    
                cleaned_values.append(cleaned_value)
                percentage_values.append(percent_value)
            else:
                # If no specific pattern is found, keep the original value
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

            # Check if the label is one of the specified ones
            if label in ["SaO² mini", "Sat. Mini", "Sat"]:
                # For these labels, add the original row without changes
                result_data.append([label] + list(data))
            else:
                # Apply the percentage extraction logic for other labels
                cleaned_data, percentage_data = PercentageExtractor.extract_percentages_from_row(data)

                # Always add the cleaned data row
                result_data.append([label] + cleaned_data)

                is_empty_row = all((not val or val == "" or (isinstance(val, float) and math.isnan(val))) for val in cleaned_data[1:])
                # If after extraction, there's percentage data but no cleaned data (excluding the label), drop the cleaned data row
                if any(percentage_data) and is_empty_row:
                    result_data.pop()  # Remove the last added (cleaned data) row

                # Add the percentage row if it contains any percentages
                if any(percentage_data):
                    result_data.append([label + "%"] + percentage_data)

        # Check if all rows have the same length
        row_lengths = [len(row) for row in result_data]
        if len(set(row_lengths)) != 1:
            print("Inconsistent row lengths found:", set(row_lengths))
        
        # Creating DataFrame from the result data
        result_df = pd.DataFrame(result_data, columns=result_data[0] if result_data else [])
        
        return result_df


    @classmethod
    def extract_percentages_from_excel(cls, directory_path, output_path=None):
        if output_path is None:
            output_path = os.path.join(os.getcwd(), 'excel_perc')

        exception_path = os.path.join(output_path, 'exception')
        os.makedirs(output_path, exist_ok=True)
        os.makedirs(exception_path, exist_ok=True)

        exception_log = []  # List to store exceptions

        for filename in filter(lambda f: f.endswith(".xlsx"), os.listdir(directory_path)):
            print(f"--------------------------Processing {filename}-------------------------- ")
            file_path = os.path.join(directory_path, filename)
            output_file_path = os.path.join(output_path, filename)

            try:
                df = pd.read_excel(file_path, header=None, engine='openpyxl')
                if df.shape[1] >= 2:
                    processed_df = cls.process_dataframe_for_percentages(df)
                    processed_df.to_excel(output_file_path, index=False, header=False, engine='xlsxwriter')
                    print(f"Processed {filename} successfully.")
                else:
                    print(f"Skipping {filename}: Insufficient columns to process.")
            except Exception as e:
                exception_msg = f"Error processing {filename}: {str(e)}"
                print(exception_msg)
                exception_log.append([filename, str(e)])

                # Move the problematic file to the exception folder
                shutil.copy(file_path, os.path.join(exception_path, filename))

        # Save the exception log to an Excel file
        if exception_log:
            exception_df = pd.DataFrame(exception_log, columns=['Filename', 'Exception'])
            exception_log_path = os.path.join(exception_path, 'exception_log.xlsx')
            exception_df.to_excel(exception_log_path, index=False)
