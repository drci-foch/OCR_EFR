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
    def extract_relative_date(reference_date, s):
        """
        Extracts a relative date based on a reference date and a string indicating the offset.
        Handles variations like "1 an plus tard", "6 mois", "2 ans", "4 jours après", etc.
        """
        num = int(re.search(r'(\d+)', s).group(1))
        if "jour" in s or "jours" in s:
            new_date = reference_date + pd.DateOffset(days=num)
        elif "mois" in s:
            new_date = reference_date + pd.DateOffset(months=num)
        elif "an" in s or "ans" in s:
            new_date = reference_date + pd.DateOffset(years=num)
        else:
            new_date = reference_date
        return new_date

    @staticmethod
    def flexible_date_formatting(date_series):
        """
        Attempt to format a series of dates using a more simplified approach.
        Handles relative date terms and relies on 'dayfirst' argument to handle various day/month formats.
        """
        cleaned_dates = []
        for idx, date_str in enumerate(date_series):
            date_str = str(date_str)
            if any(term in date_str for term in ["plus tard", "après", "mois", "an", "ans", "jour", "jours"]):
                new_date = DateFormatter.extract_relative_date(
                    cleaned_dates[-1], date_str)
                cleaned_dates.append(new_date)
            else:
                date = DateFormatter.extract_date_from_string(date_str)
                try:
                    cleaned_dates.append(pd.to_datetime(date, dayfirst=True))
                except:
                    cleaned_dates.append(date)
        return pd.Series(cleaned_dates).dt.strftime('%d/%m/%y')

    @staticmethod
    def format_dates_in_worksheet(df):
        """Formats the dates in the first row of a DataFrame using a simplified approach."""
        formatted_dates = DateFormatter.flexible_date_formatting(
            df.iloc[0, 1:])
        df.iloc[0, 1:len(formatted_dates)+1] = formatted_dates
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
                output_file_path = os.path.join(
                    directory, filename.replace(".xlsx", "_cleaned.xlsx"))
                with pd.ExcelWriter(output_file_path) as writer:
                    for sheet_name, data in all_sheets.items():
                        data.to_excel(
                            writer, sheet_name=sheet_name, index=False)
                print(
                    f"Processed: {filename} -> {filename.replace('.xlsx', '_cleaned.xlsx')}")

    @staticmethod
    def extract_relative_date(reference_date, s):
        """
        Extracts a relative date based on a reference date and a string indicating the offset.
        Handles variations like "1 an plus tard", "6 mois", "2 ans", "4 jours après", etc.
        """
        match = re.search(r'(\d+)', s)
        if match:
            num = int(match.group(1))
        else:
            # Setting num to a default value (can be changed as per requirements)
            num = 0

        if "jour" in s:
            new_date = reference_date + pd.DateOffset(days=num)
        elif "mois" in s:
            new_date = reference_date + pd.DateOffset(months=num)
        elif "an" in s or "ans" in s:
            new_date = reference_date + pd.DateOffset(years=num)
        else:
            new_date = reference_date
        return new_date
