import os
import pandas as pd
import re

class DateFormatter:

    @staticmethod
    def extract_date_from_string(s):
        """
        Extracts a date-like pattern from a string.
        """
        if pd.isna(s):
            return s
        date_pattern = re.compile(r'(\d{1,2}[./]\d{1,2}[./]\d{2})')
        match = date_pattern.search(s)
        return match.group(1) if match else s

    @staticmethod
    def flexible_date_formatting(date_series):
        """
        Attempt to format a series of dates using a more simplified approach.
        Rely on 'dayfirst' argument to handle various day/month formats.
        """
        date_series = date_series.map(lambda x: DateFormatter.extract_date_from_string(str(x)))
        date_series = pd.to_datetime(date_series, errors='coerce', dayfirst=True)
        return date_series.dt.strftime('%d/%m/%y')

    @staticmethod
    def format_dates_in_worksheet(df):
        """Formats the dates in the first row of a DataFrame using a simplified approach."""
        formatted_dates = DateFormatter.flexible_date_formatting(df.iloc[0, 1:])
        df.iloc[0, 1:] = formatted_dates
        return df

    @classmethod
    def format_dates_in_excel(cls, directory):
        """
        Format the dates in all Excel files in a specified directory.
        """
        for filename in os.listdir(directory):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(directory, filename)
                
                with pd.ExcelFile(file_path) as xls:
                    all_sheets = {}
                    
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name)
                        cleaned_df = cls.format_dates_in_worksheet(df)
                        all_sheets[sheet_name] = cleaned_df
                        
                output_file_path = os.path.join(directory, filename.replace(".xlsx", "_cleaned.xlsx"))
                with pd.ExcelWriter(output_file_path) as writer:
                    for sheet_name, data in all_sheets.items():
                        data.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"Processed: {filename} -> {filename.replace('.xlsx', '_cleaned.xlsx')}")

