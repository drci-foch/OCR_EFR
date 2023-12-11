
import os
import pandas as pd
import re
from datetime import datetime, timedelta
import calendar

class DateFormatter:
    
    french_month_to_number = {
    'jan': '01', 'janvier': '01', 'janv': '01',
    'fév': '02', 'février': '02', 'fev': '02',
    'mar': '03', 'mars': '03', 
    'avr': '04', 'avril': '04', 
    'mai': '05', 
    'jui': '06', 'juin': '06', 
    'jul': '07', 'juillet': '07', 
    'aoû': '08', 'août': '08', 'aou': '08',
    'sep': '09', 'septembre': '09', 'sept': '09',
    'oct': '10', 'octobre': '10', 
    'nov': '11', 'novembre': '11', 
    'déc': '12', 'décembre': '12', 'dec': '12',
}


    def extract_date_from_string(self, s):
        """Extracts a date-like pattern from a string and handles two-digit years using a pivot year approach."""
        if pd.isna(s) or s=="nan":
            return s
        elif '%' in s:  # Check for the presence of a percentage symbol
            return 'contains_percent'  # Return a specific flag
        else:
            # First check for month-year pattern
            if self.month_year_pattern_in_string(s):
                return self.extract_month_year_date(s)
            else:
                # This pattern matches dates in the format 11/06/14 or 11.06.14 or 11.06.2014...
                # Pattern matching various date formats
                date_patterns = [
                    re.compile(r'(\d{1,2})[./ -]+(\d{1,2})[./ -]+(\d{2,4})'),  # e.g., 18/11/08, 8-10-03
                    re.compile(r'(\d{1,2})[ ./-](\d{1,2})[ ./-](\d{2,4})'), # Updated pattern for "08 01 2015"
                    re.compile(r'(\d{1,2})[./ -]+(\d{2})'),  # e.g., 3.9, 07.98
                    re.compile(r'\b(\d{4})\b') # e.g., 2017 (as the last resort)
                ]
                not_date_patterns = [
                    re.compile(r'')
                ]
                for pattern in date_patterns:
                    match = pattern.search(s)
                    if match:
                        groups = match.groups()
                        if len(groups) == 3:
                            day, month, year = groups
                            year = self.handle_year(year)
                            print(groups)
                            return f"{day.zfill(2)}/{month.zfill(2)}/{year}"
                        elif len(groups) == 2:
                            month, year = groups
                            year = self.handle_year(year)
                            print(groups)
                            return f"01/{month.zfill(2)}/{year}"
                        elif len(groups) == 1:
                            # If only year is found
                            year = groups[0]
                            print(groups)
                            return f"01/01/{year}"

                        else:
                            # This pattern matches dates in the format 01042014 for 01/04/2014
                            date_pattern_strict = re.compile(r'(\d{1,2})(\d{1,2})(\d{2,4})')
                            match_strict = date_pattern_strict.search(s)
                            if match_strict:
                                day, month, year = match_strict.groups()
                                year_corrected = self.handle_year(year)
                                return  f"{day.zfill(2)}/{month.zfill(2)}/{year_corrected}"

            return s

    def handle_year(self, year_str):
        """Handle two-digit years using a pivot year approach."""
        if len(year_str) == 2:  # Apply pivot year logic only for two-digit years
            pivot_year = 30  # Years up to 30 are considered as 2000-2030
            century = '19' if int(year_str) > pivot_year else '20'
            year_str = century + year_str

        return year_str.zfill(4)  # Ensure year is always four digits



    def extract_relative_date(self, reference_date, s):
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

    def month_year_pattern_in_string(self, s):
        month_year_pattern = re.compile(r'((\w+)(\s+)(\d{4}))', re.IGNORECASE)
        return bool(month_year_pattern.search(s))

    def split_and_distribute(self, table):
        for i, cell in enumerate(table):
            # Skip header row
            if i == 0:
                continue

            if cell and len(cell.split()) > 1:
                # Count adjacent null cells
                null_count = 0
                for j in range(i + 1, len(table)):
                    if pd.isna(table[j]) or table[j].strip() == '':
                        null_count += 1
                    else:
                        break

                # Split the cell content if the number of elements matches the number of adjacent nulls + 1
                split_content = cell.split()
                if len(split_content) == null_count + 1:
                    for k, content in enumerate(split_content):
                        table[i + k] = content

        return table


    def extract_month_year_date(self, s):
        if pd.isna(s):
            return s

        # Split the input string into words
        words = s.split()

        # Initialize lists to store corrected month and year
        corrected_month = None
        corrected_year = None

        # Function to calculate similarity between two strings
        def string_similarity(s1, s2):
            s1 = s1.lower()
            s2 = s2.lower()
            n = len(s1)
            m = len(s2)
            dp = [[0] * (m + 1) for _ in range(n + 1)]
            for i in range(n + 1):
                for j in range(m + 1):
                    if i == 0:
                        dp[i][j] = j
                    elif j == 0:
                        dp[i][j] = i
                    elif s1[i - 1] == s2[j - 1]:
                        dp[i][j] = dp[i - 1][j - 1]
                    else:
                        dp[i][j] = 1 + min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1])
            return dp[n][m]

        # Iterate through words in the input string
        for word in words:
            # Initialize variables to store best match and minimum similarity score
            best_match = None
            min_similarity = float('inf')

            # Try to find the best matching valid month name
            for month_name, month_number in self.french_month_to_number.items():
                similarity = string_similarity(word, month_name)
                if similarity < min_similarity:
                    min_similarity = similarity
                    best_match = month_name

            # If a valid month name is found, correct it
            if min_similarity <= 2:  # You can adjust the similarity threshold as needed
                corrected_month = self.french_month_to_number[best_match]
            else:
                # Check if the word is a valid year
                if word.isdigit() and len(word) == 4:
                    corrected_year = self.handle_year(word)

        # If both month and year are found, construct the formatted date
        if corrected_month and corrected_year:
            formatted_date = f"01/{corrected_month}/{corrected_year}"
            return formatted_date
        else:
            return None


    def adjust_date(self, date_to_adjust, reference_date):
        """Adjusts a date to make it consistent with a reference date."""
        year = reference_date.year if date_to_adjust.year < reference_date.year else date_to_adjust.year
        month = reference_date.month if date_to_adjust.month < reference_date.month else date_to_adjust.month

        # Find the last day of the new month
        last_day_of_month = calendar.monthrange(year, month)[1]

        # Adjust the day if necessary
        day = min(date_to_adjust.day, last_day_of_month)

        # Construct the adjusted date
        return pd.Timestamp(year=year, month=month, day=day)

    def correct_date_chronology(self, date_series):
        corrected_dates = []
        last_valid_date = None

        for date in date_series:
            current_date = pd.to_datetime(date, format='%d/%m/%Y', errors='coerce')

            if pd.isnull(current_date):
                corrected_dates.append(None)
                continue

            if last_valid_date and current_date < last_valid_date:
                # Use adjust_date to correct the current_date
                current_date = self.adjust_date(current_date, last_valid_date)

            corrected_dates.append(current_date)
            last_valid_date = current_date

        # Convert datetime objects back to string format if needed
        return pd.Series([date.strftime('%d/%m/%Y') if pd.notnull(date) else None for date in corrected_dates])


    def flexible_date_formatting(self, series, reference_date=None):
        def truncate_date(date):
            if date is None:
                return date

            # Check if the 'date' matches the YYYYMMDD pattern
            match = re.match(r'(\d{2})(\d{2})(\d{4})', date)
            if match:
                day, month, year = match.groups()
            else:
                day, month, year = date.split('/')

            # Truncate month to 12 if it is greater than 12
            month = '12' if int(month) > 12 else month
            # Truncate day to last day of month if it is too large
            day = str(min(int(day), 31)).zfill(2)
            year = '20' + year if len(year) == 2 else year

            truncated_date = f"{day}/{month}/{year}"
            return truncated_date

        cleaned_dates = []
        for i, date in enumerate(series):
            try:
                if pd.isna(date):
                    cleaned_date = "date" if i == 0 else None
                else:
                    date_str = str(date)

                cleaned_date = self.extract_date_from_string(str(date))

                # Apply truncation
                truncated_date = truncate_date(cleaned_date)

                # Check if truncated_date is None
                if truncated_date is None:
                    final_date = None  # Or set a default value or placeholder
                else:
                    # Parse and format the date
                    final_date = datetime.strptime(truncated_date, '%d/%m/%Y').strftime('%d/%m/%Y')

                cleaned_dates.append(final_date)

            except ValueError as e:
                print(f"Invalid date found: {date}, error: {e}, replacing with original date")
                cleaned_dates.append(str(date))

        return pd.Series(cleaned_dates)

    def format_dates_in_excel(self, directory_path: str, output_path = "C:\\Users\\benysar\\Desktop\\Github\\OCR_EFR\\QuickScanEFR\\pdf_TextMachina\\excel_date", problem_directory="C:\\Users\\benysar\\Desktop\\Github\\OCR_EFR\\QuickScanEFR\\pdf_TextMachina\\problem"):
        for filename in os.listdir(directory_path):
            if filename.endswith(".xlsx"):
                filepath = os.path.join(directory_path, filename)
                filepath_out =  os.path.join(output_path, filename)
                try:
                    print(f"Processing file: {filepath}")

                    df = pd.read_excel(filepath, header=None)
                    #Drop columns and lines that are completely empty
                    df = df.dropna(axis=1, how='all')
                    df = df.dropna(axis=0, how='all')

                    # Step 1: Find the index of the first non-null element in column 0, except for the first row
                    first_non_null_index = df.iloc[1:, 0].notnull().idxmax()

                    # Step 2: Get the index of the row just before the first non-null row
                    desired_index = first_non_null_index - 1

                    # Step 3: Retrieve the row
                    first_row = df.iloc[desired_index]

                    #print(f"Original first row:\n{first_row}")
                    corrected_first_row = self.flexible_date_formatting(first_row)
                    
                    contains_percent = corrected_first_row.apply(lambda x: 'contains_percent' if '%' in str(x) else x)

                    if 'contains_percent' in contains_percent.values:
                        # Try the row above
                        desired_index -= 1
                        first_row = df.iloc[desired_index]
                        corrected_first_row = self.flexible_date_formatting(first_row)
                        contains_percent = corrected_first_row.apply(lambda x: 'contains_percent' if '%' in str(x) else x)
                        if 'contains_percent' in contains_percent.values:
                            # If still contains percent, move file to problem folder
                            problem_file_path = os.path.join(problem_directory, filename)
                            if os.path.exists(problem_file_path):
                                os.remove(problem_file_path)
                            os.rename(filepath, problem_file_path)
                            continue  # Skip the rest of the processing for this file

                    corrected_first_row = self.split_and_distribute(corrected_first_row)
                    corrected_first_row = pd.to_datetime(corrected_first_row, errors='coerce', format='%d/%m/%Y')

                    if isinstance(corrected_first_row, pd.Series):
                        corrected_first_row = corrected_first_row.dt.strftime('%d/%m/%Y')
                    elif isinstance(corrected_first_row, pd.DatetimeIndex):
                        corrected_first_row = corrected_first_row.strftime('%d/%m/%Y')

                    #print(f"Corrected first row:\n{corrected_first_row}")

                    df.iloc[0] = self.correct_date_chronology(corrected_first_row)
                    #print(f"Correct date chronology:\n{df.iloc[0]}")

                    df.to_excel(filepath_out, index=False, header=False)
                    
                except ValueError as e:
                    if "could not broadcast input array from shape" in str(e) or "Must have equal len keys and value when setting with an iterable" in str(e):
                        problem_file_path = os.path.join(problem_directory, filename)
                        # Check if the file exists in the problem directory and remove it if it does
                        if os.path.exists(problem_file_path):
                            os.remove(problem_file_path)
                        os.rename(filepath, problem_file_path)
                        print(f"Problem with file {filename}. Moved to 'problem' folder due to error: {e}")
                    else:
                        # Handle other ValueError cases or re-raise the exception
                        raise
