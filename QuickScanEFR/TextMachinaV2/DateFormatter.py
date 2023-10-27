import os
import re
import pandas as pd
from excel_utils import load_excel_workbook, get_processed_file_name, save_excel_data

# Refactoring the DateFormatter.py module

class DateFormatterRefactored:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.data = None

    def extract_date_from_string(self, s):
        """Extract a date-like pattern from a string."""
        date_pattern = r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})'
        matches = re.findall(date_pattern, s)
        return matches[0] if matches else None

    def extract_relative_date(self, reference_date, offset_string):
        """Calculate a date based on a reference date and an offset string."""
        match = re.search(r'(\d+)', offset_string)
        if match:
            num = int(match.group(1))
        else:
            # Setting num to a default value (can be changed as per requirements)
            num = 0

        if "jour" in offset_string:
            new_date = reference_date + pd.DateOffset(days=num)
        elif "mois" in offset_string:
            new_date = reference_date + pd.DateOffset(months=num)
        elif "an" in offset_string or "ans" in offset_string:
            new_date = reference_date + pd.DateOffset(years=num)
        else:
            new_date = reference_date
        return new_date

    def correct_date_chronology(self, date_series):
        """Ensure the chronological order of dates in a series and correct any inconsistencies."""
        dates = [pd.to_datetime(date, dayfirst=True) for date in date_series]
        for i in range(1, len(dates) - 1):
            prev_date = dates[i - 1]
            curr_date = dates[i]
            next_date = dates[i + 1]
            if not prev_date <= curr_date <= next_date:
                if prev_date.year < curr_date.year < next_date.year:
                    pass
                elif curr_date.year < prev_date.year:
                    curr_date = curr_date.replace(year=prev_date.year)
                else:
                    curr_date = curr_date.replace(year=next_date.year)
                if prev_date.month < curr_date.month < next_date.month:
                    pass
                elif curr_date.month < prev_date.month:
                    curr_date = curr_date.replace(month=prev_date.month)
                else:
                    curr_date = curr_date.replace(month=next_date.month)
                if prev_date.day < curr_date.day < next_date.day:
                    pass
                elif curr_date.day < prev_date.day:
                    curr_date = curr_date.replace(day=prev_date.day)
                else:
                    curr_date = curr_date.replace(day=next_date.day)
                dates[i] = curr_date
        return pd.Series(dates).dt.strftime('%d/%m/%y')


    def delete_empty_date_columns(self, df):
        """Remove columns that only contain NaN values (empty columns) from a DataFrame."""
        return df.dropna(axis=1, how='all', thresh=2)  # Keep columns with at least 2 non-NaN values

    def format_dates_in_worksheet(self, df):
        """Format the dates in the first row of a DataFrame."""
        formatted_dates = DateFormatterRefactored.flexible_date_formatting(
            df.iloc[0, 1:])
        df.iloc[0, 1:len(formatted_dates)+1] = formatted_dates
        df = DateFormatterRefactored.delete_empty_date_columns(df)
        return df

    def format_dates_in_excel(self, directory_path):
        """Format the dates in all Excel files in a specified directory."""
        for filename in os.listdir(directory_path):
            if filename.endswith("_modified.xlsx"):
                file_path = os.path.join(directory_path, filename)
                self.data = load_excel_workbook(file_path)
                processed_data = self.format_dates_in_worksheet(self.data)
                save_path = get_processed_file_name(file_path, "corrected")
                save_excel_data(processed_data, save_path)
                print(f"Processed {file_path} and saved as {save_path}")

