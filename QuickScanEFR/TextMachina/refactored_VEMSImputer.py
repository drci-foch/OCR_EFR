import os

import pandas as pd


class VEMSCorrector:
    def __init__(self, excel_file_path=None):
        """
        Initialize the VEMSCorrector.

        Parameters:
        - excel_file_path (str, optional): Path to the input Excel file.
        """
        self.excel_file_path = excel_file_path
        self.workbook = None
        self.sheet_names = []
        self.combined_data = None
        if excel_file_path:
            self._load_workbook()
            
    def _is_percent_value_usable(self, value):
        """
        Check if a given percentage value is usable.

        Parameters:
        - value (str): The percentage value to check.

        Returns:
        - bool: True if the value is usable, False otherwise.
        """
        try:
            float(value.replace('%', ''))
            return True
        except:
            return False
        
    def _load_workbook(self):
        """
        Load workbook and initialize necessary attributes.
        """
        self.combined_data = pd.read_excel(self.excel_file_path, header=None)

    def compute_missing_vems(self):
        # Extracting rows related to VEMS and VEMS% from the combined data
        vems_row = self.combined_data[self.combined_data[0] == 'VEMS']
        vems_percent_row = self.combined_data[self.combined_data[0] == 'VEMS%']

        if not vems_row.empty and vems_row.shape[1] > 1:
            dates = self.combined_data.columns[1:].to_list()
            vems_values = vems_row.iloc[0, 1:].values
            vems_percent_values = [
                float(str(val).replace('%', '')) / 100 if pd.notnull(val) and self._is_percent_value_usable(str(val))
                else None for val in vems_percent_row.iloc[0, 1:].values
            ]

            for idx, vems_percent in enumerate(vems_percent_values):
                if vems_values[idx] is None and vems_percent is not None:
                    left_vems = None
                    right_vems = None
                    left_date_diff = None
                    right_date_diff = None

                    # Calculate the date for the current missing value
                    current_date = pd.to_datetime(dates[idx])

                    # Scan leftward for the nearest non-null VEMS value
                    for left_idx in range(idx - 1, -1, -1):
                        if vems_values[left_idx] is not None:
                            left_vems = vems_values[left_idx]
                            left_date = pd.to_datetime(dates[left_idx])
                            left_date_diff = (current_date - left_date).days / 30.44  # Average days per month
                            break

                    # Scan rightward for the nearest non-null VEMS value
                    for right_idx in range(idx + 1, len(vems_values)):
                        if vems_values[right_idx] is not None:
                            right_vems = vems_values[right_idx]
                            right_date = pd.to_datetime(dates[right_idx])
                            right_date_diff = (right_date - current_date).days / 30.44
                            break

                    # Impute missing VEMS value based on nearest non-null values and date differences
                    if ((left_vems is not None and left_date_diff < 12) or
                        (right_vems is not None and right_date_diff < 12)):
                        
                        if left_vems is not None and right_vems is not None:
                            vems_estimated = (left_vems + right_vems) / 2
                        elif left_vems is not None:
                            vems_estimated = left_vems
   
                        elif right_vems is not None:
                            vems_estimated = right_vems
                            
                        else:
                            vems_estimated = None

                        if vems_estimated is not None:
                            vems_values[idx] = vems_estimated
                            print(vems_estimated)

            # Update the DataFrame with the new VEMS values
            self.combined_data.iloc[vems_row.index[0], 1:] = vems_values


        else:
            print("Keeping original DataFrame: vems_row is empty or does not have enough columns.")


    # def compute_missing_vems_percent(self):
    #     """
    #     Compute missing VEMS% values based on provided VEMS values.
    #     """
    #     # Extracting rows related to VEMS and VEMS% from the combined data
    #     vems_row = self.combined_data[self.combined_data[0] == 'VEMS']
    #     vems_percent_row = self.combined_data[self.combined_data[0] == 'VEMS%']

    #     # Extract VEMS, VEMS%, and date values
    #     dates = self.combined_data.columns[1:].to_list()
    #     vems_values = vems_row.iloc[0, 1:].values
    #     vems_percent_values = [
    #         float(str(val).replace('%', '')) / 100 if pd.notnull(val) and self._is_percent_value_usable(str(val))
    #         else None for val in vems_percent_row.iloc[0, 1:].values
    #     ]

    #     # Placeholder for extrapolated VEMS% values
    #     extrapolated_vems_percent = []

    #     # Loop through each date
    #     for idx, (vems, vems_percent) in enumerate(zip(vems_values, vems_percent_values)):
    #         vems = float(vems.replace('%', '')) if pd.notnull(vems) else None
            
    #         if pd.notnull(vems) and vems_percent is None:
    #             # If VEMS% is missing but VEMS is provided

    #             # Find closest non-null VEMS% values on both sides
    #             left_idx = idx - 1
    #             while left_idx >= 0 and vems_percent_values[left_idx] is None:
    #                 left_idx -= 1

    #             right_idx = idx + 1
    #             while right_idx < len(vems_percent_values) and vems_percent_values[right_idx] is None:
    #                 right_idx += 1

    #             # Calculate difference in months for the dates
    #             current_date = pd.to_datetime(dates[idx])
    #             if 0 <= left_idx < len(vems_percent_values):
    #                 left_date = pd.to_datetime(dates[left_idx])
    #                 left_diff_months = (
    #                     current_date.year - left_date.year) * 12 + current_date.month - left_date.month
    #             else:
    #                 left_diff_months = None

    #             if 0 <= right_idx < len(vems_percent_values):
    #                 right_date = pd.to_datetime(dates[right_idx])
    #                 right_diff_months = (
    #                     right_date.year - current_date.year) * 12 + right_date.month - current_date.month
    #             else:
    #                 right_diff_months = None

    #             # If either left or right VEMS% value is found and its date is less than 12 months apart, extrapolate VEMS%
    #             if (0 <= left_idx < len(vems_percent_values) and (left_diff_months is None or left_diff_months < 12)) or (0 <= right_idx < len(vems_percent_values) and (right_diff_months is None or right_diff_months < 12)):
    #                 if left_diff_months is not None and right_diff_months is not None:
    #                     vems_percent_estimated = ((vems / float(vems_values[left_idx]) * vems_percent_values[left_idx]) +
    #                                             (vems / float(vems_values[right_idx]) * vems_percent_values[right_idx])) / 2
    #                 elif left_diff_months is not None:
    #                     vems_percent_estimated = vems / float(vems_values[left_idx]) * vems_percent_values[left_idx]
    #                 else:
    #                     vems_percent_estimated = vems / float(vems_values[right_idx]) * vems_percent_values[right_idx]

    #                 extrapolated_vems_percent.append(
    #                     (idx, round(vems_percent_estimated * 100, 2)))  # Convert back to percentage
    #             else:
    #                 extrapolated_vems_percent.append((idx, None))
    #         else:
    #             extrapolated_vems_percent.append((idx, None))

    #     # Update VEMS% row with extrapolated values
    #     for idx, value in extrapolated_vems_percent:
    #         if value is not None:
    #             vems_percent_values[idx] = f"{value}%"

    #     # Update the dataframe with the new VEMS% values
    #     self.combined_data.iloc[vems_percent_row.index[0], 1:] = vems_percent_values


    def save_combined_data(self, output_file_path=None):
        """
        Save the combined data with computed VEMS values to an Excel file.

        Parameters:
        - output_file_path (str, optional): Path for saving the combined data.
        """
        # Define the new folder name
        new_folder = 'excel_correctedVEMS'

        # Extract the directory from the existing file path
        dir_path = os.path.dirname(self.excel_file_path)

        # Create the new directory path
        new_dir_path = os.path.join(dir_path, new_folder)

        # Create the new folder if it doesn't exist
        if not os.path.exists(new_dir_path):
            os.makedirs(new_dir_path)

        # Construct the new file path
        if output_file_path is None:
            output_file_path = os.path.join(new_dir_path, os.path.basename(self.excel_file_path))

        # Save the file
        self.combined_data.to_excel(output_file_path, index=False, header=False,engine='xlsxwriter')


    def correct_multiple_docs(self, directory_path):
        excel_files = [f for f in os.listdir(directory_path) if f.endswith('.xlsx')]
        exception_log = []  # List to store exceptions

        for excel_file in excel_files:
            try:
                print(f'----------------Processing {excel_file}----------------')
                self.excel_file_path = os.path.join(directory_path, excel_file)
                self._load_workbook()
                self.compute_missing_vems()
                # self.compute_missing_vems_percent()
                self.save_combined_data()

            except Exception as e:
                print(f"Error processing {excel_file}: {e}")
                exception_log.append((excel_file, str(e)))
                # Copy the file to the 'problem' sub-folder
                problem_folder = os.path.join(directory_path, 'problem')
                if not os.path.exists(problem_folder):
                    os.makedirs(problem_folder)
                problem_file_path = os.path.join(problem_folder, excel_file)
                os.rename(os.path.join(directory_path, excel_file), problem_file_path)

        # Save the exception log to an Excel file
        if exception_log:
            problem_log_df = pd.DataFrame(exception_log, columns=['File Name', 'Exception'])
            problem_log_path = os.path.join(directory_path, 'problem', 'exception_log.xlsx')
            problem_log_df.to_excel(problem_log_path, index=False)