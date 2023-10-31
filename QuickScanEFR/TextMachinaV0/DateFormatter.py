import os
import pandas as pd
import re
from dateutil import parser

class DateFormatter:
    
    french_month_to_english = {
        'janvier': 'January',
        'février': 'February',
        'mars': 'March',
        'avril': 'April',
        'mai': 'May',
        'juin': 'June',
        'juillet': 'July',
        'août': 'August',
        'septembre': 'September',
        'octobre': 'October',
        'novembre': 'November',
        'décembre': 'December'
    }

    @staticmethod
    def extract_date_from_string(s):
        """Extracts a date-like pattern from a string."""
        if pd.isna(s):
            return s
        date_pattern = re.compile(r'(\d{1,2}[./]\d{1,2}[./]\d{2})')
        match = date_pattern.search(s)
        return match.group(1) if match else s

    @staticmethod
    def extract_relative_date(reference_date, s):
        """Extracts a relative date based on a reference date and a string indicating the offset."""
        try:
            match = re.search(r'(\d+)', s)
            if match:
                num = int(match.group(1))
            else:
                # Setting num to a default value (can be changed as per requirements)
                num = 0
            
            if "jour" in s:
                new_date = reference_date + pd.DateOffset(days=num)
            elif "J" in s.upper():
                new_date = reference_date + pd.DateOffset(days=num)
            elif "mois" in s:
                new_date = reference_date + pd.DateOffset(months=num)
            elif "M" in s.upper():
                new_date = reference_date + pd.DateOffset(months=num)
            elif "an" in s or "ans" in s:
                new_date = reference_date + pd.DateOffset(years=num)
            else:
                new_date = reference_date
            return new_date
        except Exception as e:
            return s

    @staticmethod
    def extract_month_year_date(s):
        """Extracts a date in the format 'Month Year' and converts it to a valid date format."""
        try:
            s_lower = s.lower()  # Convert the input string to lowercase
            for french_month, english_month in DateFormatter.french_month_to_english.items():
                s_lower = s_lower.replace(french_month.lower(), english_month.lower())
            
            date = parser.parse(s_lower, fuzzy=True)
            formatted_date = date.strftime('%d/%m/%y')
            return formatted_date
        except Exception as e:
            return s

    @staticmethod
    def correct_date_chronology(date_series):
        """Ensure the chronology of dates and correct any inconsistencies."""
        try:
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
        except Exception as e:
            return date_series

    @staticmethod
    def flexible_date_formatting(date_series):
        """Attempt to format a series of dates using a simplified approach."""
        cleaned_dates = []
        for idx, date_str in enumerate(date_series):
            date_str = str(date_str)
            if any(term in date_str for term in ["plus tard", "après", "mois", "an", "ans", "jour", "jours"]):
                new_date = DateFormatter.extract_relative_date(
                    cleaned_dates[-1], date_str)
                cleaned_dates.append(new_date)
            else:
                date = DateFormatter.extract_date_from_string(date_str)
                if date != date_str:
                    cleaned_dates.append(pd.to_datetime(date))
                else:
                    month_year_date = DateFormatter.extract_month_year_date(date_str)
                    if month_year_date != date_str:
                        cleaned_dates.append(pd.to_datetime(month_year_date))
                    else:
                        cleaned_dates.append(date_str)
        return DateFormatter.correct_date_chronology(cleaned_dates)

    @staticmethod
    def flexible_date_formatting(date_series):
        """Attempt to format a series of dates using a simplified approach."""
        cleaned_dates = []
        for idx, date_str in enumerate(date_series):
            date_str = str(date_str)
            if any(term in date_str for term in ["plus tard", "après", "mois", "an", "ans", "jour", "jours"]):
                new_date = DateFormatter.extract_relative_date(
                    cleaned_dates[-1], date_str)
                cleaned_dates.append(new_date)
            else:
                date = DateFormatter.extract_date_from_string(date_str)
                if date != date_str:
                    cleaned_dates.append(pd.to_datetime(date))
                else:
                    month_year_date = DateFormatter.extract_month_year_date(date_str)
                    if month_year_date != date_str:
                        cleaned_dates.append(pd.to_datetime(month_year_date))
                    else:
                        # Handle missing or incomprehensible dates by inserting them between adjacent valid dates
                        if cleaned_dates:
                            prev_date = cleaned_dates[-1]
                            next_date = None
                            if idx + 1 < len(date_series):
                                next_date_str = str(date_series[idx + 1])
                                next_date = pd.to_datetime(next_date_str)
                            
                            if prev_date and next_date:
                                inserted_date = prev_date + pd.DateOffset(days=1)  # You can adjust this logic as needed
                                cleaned_dates.append(inserted_date)
                            else:
                                cleaned_dates.append(date_str)
                        else:
                            cleaned_dates.append(date_str)
        return DateFormatter.correct_date_chronology(cleaned_dates)

    @staticmethod
    def delete_empty_date_columns(df):
        """Delete empty columns containing only dates."""
        try:
            # Check if the entire column has NaN values in rows other than the first row
            is_empty_column = df.apply(lambda col: pd.isna(col.iloc[1:]).all())

            # Find the columns where the first row is not NaN
            non_empty_columns = df.columns[~is_empty_column]
            return df[non_empty_columns]
        except Exception as e:
            return df

    @staticmethod
    def format_dates_in_worksheet(df):
        """Formats the dates in the first row of a DataFrame using a simplified approach."""
        try:
            formatted_dates = DateFormatter.flexible_date_formatting(
                df.iloc[0, 1:])
            df.iloc[0, 1:len(formatted_dates)+1] = formatted_dates
            df = DateFormatter.delete_empty_date_columns(df)
            return df
        except Exception as e:
            return df

    @classmethod
    def format_dates_in_excel(cls, directory):
        """Format the dates in all Excel files in a specified directory."""
        try:
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
        except Exception as e:
            print(f"Error processing {filename}: {e}")
