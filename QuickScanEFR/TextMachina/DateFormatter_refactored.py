
import os
import pandas as pd
import re
from datetime import datetime, timedelta

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
    def extract_date_from_string(self, s):
        """Extracts a date-like pattern from a string."""
        if pd.isna(s):
            return s
        date_pattern = re.compile(r'(\d{1,2}[./]\d{1,2}[./]\d{2})')
        match = date_pattern.search(s)
        return match.group(1) if match else s


    def extract_relative_date(self, string, reference_date):
        """Extracts a relative date based on a reference date."""
        ref_date = datetime.strptime(reference_date, "%d %B %Y")
        if "day" in string:
            days = int(re.search(r"(\d+) day", string).group(1))
            new_date = ref_date + timedelta(days=days)
            return new_date.strftime("%d/%B/%Y")
        return string

    def extract_month_year_date(self, s):
        """Extracts a date from a string in the format 'Month Year'."""
        if pd.isna(s):
            return s
        match = re.search(r'(\w+)\s+(\d{4})', s)
        if match:
            month, year = match.groups()
            month = self.french_month_to_english.get(month.lower(), month)
            return datetime.strptime(f"{month} {year}", "%B %Y").date()
        else:
            return pd.to_datetime(s, dayfirst=True, errors='coerce')

    
    @staticmethod
    def correct_date_chronology(date_series):
        """Ensure the chronology of dates and correct any inconsistencies."""
        try:
            dates = [pd.to_datetime(date, dayfirst=True) for date in date_series]

            # Start loop from the second date and end on the second last date
            for i in range(1, len(dates) - 2):
                prev_date = dates[i - 1]
                curr_date = dates[i]
                next_date = dates[i + 1]

                if not prev_date <= curr_date <= next_date:
                    # Correct the year
                    if curr_date.year < prev_date.year or curr_date.year > next_date.year:
                        curr_date = curr_date.replace(year=prev_date.year)

                    # Correct the month
                    if curr_date.month < prev_date.month or curr_date.month > next_date.month:
                        curr_date = curr_date.replace(month=prev_date.month)

                    # Correct the day
                    if curr_date.day < prev_date.day or curr_date.day > next_date.day:
                        curr_date = curr_date.replace(day=prev_date.day)

                    dates[i] = curr_date

            return pd.Series(dates)

        except Exception as e:
            return date_series



    def flexible_date_formatting(self, date_series):
        cleaned_dates = []
        
        # Convert the date_series to a list to handle non-Series inputs
        date_list = list(date_series)
        
        for idx, date_str in enumerate(date_list):
            # Convert the date_str to a string and handle NaN or float values
            date_str = str(date_str)

            if pd.isna(date_str) or date_str == 'nan':
                cleaned_dates.append(date_str)
                continue
            
            # Check for relative dates
            if any(term in date_str for term in ["plus tard", "après", "mois", "an", "ans", "jour", "jours"]):
                new_date = self.extract_relative_date(cleaned_dates[-1], date_str)
                cleaned_dates.append(new_date)
            else:
                # Check for explicit dates
                date = self.extract_date_from_string(date_str)
                if date != date_str:
                    cleaned_dates.append(pd.to_datetime(date))
                else:
                    # Check for month-year dates
                    month_year_date = self.extract_month_year_date(date_str)
                    if month_year_date != date_str:
                        cleaned_dates.append(pd.to_datetime(month_year_date))
                    else:
                        # Handle missing or incomprehensible dates
                        if cleaned_dates:
                            prev_date = cleaned_dates[-1]
                            next_date = None
                            if idx + 1 < len(date_list):
                                next_date_str = str(date_list[idx + 1])
                                next_date = pd.to_datetime(next_date_str, errors='coerce')
                            
                            if prev_date and next_date:
                                inserted_date = pd.to_datetime(prev_date) + pd.DateOffset(days=1)
                                cleaned_dates.append(inserted_date)
                            else:
                                cleaned_dates.append(date_str)
                        else:
                            cleaned_dates.append(date_str)

        
        return self.correct_date_chronology(pd.to_datetime(cleaned_dates, format="%d/%m/%y", errors='coerce'))
    
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

    def format_dates_in_excel(self, directory_path):
        for filename in os.listdir(directory_path):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(directory_path, filename)
                
                # Read the Excel file without headers
                df = pd.read_excel(file_path, header=None)
                
                # Get the entire first row as a series
                first_row = df.iloc[0]
                
                # Apply the method on the entire series to correct the dates
                corrected_first_row = self.flexible_date_formatting(first_row)
                
                # Convert to datetime, handling any non-datetime values
                corrected_first_row = pd.to_datetime(corrected_first_row, errors='coerce')
                
                # Convert datetime objects to strings in "dd/mm/yyyy" format
                corrected_first_row = corrected_first_row.dt.strftime('%d/%m/%Y')
                
                # Replace the first row with the corrected dates
                df.iloc[0] = corrected_first_row
                
                # Save the corrected DataFrame back to the Excel file
                df.to_excel(file_path, index=False, header=False)
