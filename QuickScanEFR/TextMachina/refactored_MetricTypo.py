import os
import pandas as pd
from typing import List

class TypoCorrector:
    def __init__(self, standardized_strings: List[str]):
        self.standardized_strings = standardized_strings

    @staticmethod
    def levenshtein_distance(s1: str, s2: str) -> int:
        s1 = str(s1)  # Ensure s1 is a string
        s2 = str(s2)  # Ensure s2 is a string

        if len(str(s1)) > len(str(s2)):
            s1, s2 = s2, s1

        distances = range(len(s1) + 1)
        for index2, char2 in enumerate(s2):
            new_distances = [index2 + 1]
            for index1, char1 in enumerate(s1):
                if char1 == char2:
                    new_distances.append(distances[index1])
                else:
                    new_distances.append(1 + min((distances[index1], distances[index1 + 1], new_distances[-1])))
            distances = new_distances

        return distances[-1]
    
    def correct(self, string: str, relevant_metrics: List[str], max_distance: int = 3) -> str:
        best_match = None
        best_distance = float('inf')
        for standard in relevant_metrics:
            distance = self.levenshtein_distance(string, standard)
            if distance < best_distance:
                best_distance = distance
                best_match = standard
        return best_match if best_distance <= max_distance else " "

    def correct_multiple_docs(self, directory_path: str, directory_output="C:\\Users\\benysar\\Desktop\\Github\\OCR_EFR\\QuickScanEFR\\pdf_TextMachina\\excel_metrics", problem_directory="C:\\Users\\benysar\\Desktop\\Github\\OCR_EFR\\QuickScanEFR\\pdf_TextMachina\\problem"):
        """Correct typos in the metric names in Excel files."""
        for filename in os.listdir(directory_path):
            if filename.endswith(".xlsx"):
                filepath = os.path.join(directory_path, filename)

                try:
                    df = pd.read_excel(filepath, header=None)
                    if df.empty or df.shape[1] == 0:
                        raise ValueError("DataFrame is empty or missing first column")

                    if pd.isnull(df.iloc[0, 0]):
                        df.iloc[0, 0] = 'date'

                    metric_values = df.iloc[:, 0]
                    unique_metrics = metric_values.dropna().unique().tolist()
                    relevant_metrics = [self.correct(metric, self.standardized_strings) for metric in unique_metrics]
                    #print(filename, relevant_metrics)
                    # Create a new series with corrected values
                    corrected_series = metric_values.apply(lambda x: self.correct(str(x), relevant_metrics) if pd.notnull(x) else x)
                    print(filename)
                    print(f'corrected_series :{corrected_series}')
                    print(f'df first row :{df.iloc[:, 0]}')


                    # Check and ensure that the lengths of the series match before assignment
                    if len(corrected_series) == len(df.iloc[:, 0]):
                        df.iloc[:, 0] = corrected_series
                    else:
                        # Handle the case where the lengths do not match
                        # This could involve padding the series or truncating it
                        # For now, let's just print out a message
                        print(f"Length mismatch in file {filename}: {len(corrected_series)} vs {len(df.iloc[:, 0])}")
                        # Optionally, you could raise an error here or handle it as needed

                    # Create the output file path based on the directory_output and filename
                    output_filepath = os.path.join(directory_output, filename)

                    # Save the corrected DataFrame to Excel with a valid file extension
                    with pd.ExcelWriter(output_filepath, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, header=False)

                except (IndexError, ValueError) as e:
                    problem_file_path = os.path.join(problem_directory, filename)

                    # Check if the file already exists in the problem folder
                    if os.path.exists(problem_file_path):
                        # Delete the existing file
                        os.remove(problem_file_path)

                    # Rename the file to move it to the problem folder
                    os.rename(filepath, problem_file_path)
                    print(f"Problem with file {filename}. Moved to 'problem' folder due to error: {e}")
