import os
import pandas as pd
from typing import List

class TypoCorrector:
    def __init__(self, standardized_strings: List[str]):
        self.standardized_strings = standardized_strings

    @staticmethod
    def levenshtein_distance(s1: str, s2: str) -> int:
        s1 = str(s1)
        s2 = str(s2)

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

    def correct_data_in_memory(self, data_dict):
        """Process data that's in memory with detailed logging"""
        corrected_data = {}
        print("\nStarting metric typo correction:")
        
        for filename, df in data_dict.items():
            print(f"\nProcessing file: {filename}")
            try:
                df = df.copy()
                if df.empty or df.shape[1] == 0:
                    raise ValueError("DataFrame is empty or missing first column")

                # Get the metrics column (either index or first column)
                if df.index.name is None:
                    metrics_col = df.columns[0]
                    metric_values = df[metrics_col]
                else:
                    metrics_col = df.index.name or 'index'
                    metric_values = df.index

                if pd.isnull(metric_values.iloc[0]):
                    metric_values.iloc[0] = 'date'
                    print("Set null first value to 'date'")

                unique_metrics = metric_values.dropna().unique().tolist()
                print("\nOriginal unique metrics:", unique_metrics)
                
                relevant_metrics = [self.correct(metric, self.standardized_strings) for metric in unique_metrics]
                print("Corrected unique metrics:", relevant_metrics)

                # Create a dictionary to track changes
                changes = {}
                
                def correct_and_log(x):
                    if pd.isnull(x):
                        return x
                    corrected = self.correct(str(x), relevant_metrics)
                    if str(x) != corrected and corrected != " ":
                        changes[str(x)] = corrected
                    return corrected

                corrected_series = metric_values.apply(correct_and_log)

                # Print the changes
                if changes:
                    print("\nMetric corrections made:")
                    for original, corrected in changes.items():
                        print(f"  '{original}' -> '{corrected}'")
                else:
                    print("\nNo corrections needed")

                # Apply corrections
                if df.index.name is None:
                    df[metrics_col] = corrected_series
                else:
                    df.index = corrected_series

                corrected_data[filename] = df
                print(f"Successfully processed {filename}")

            except Exception as e:
                print(f"Problem with DataFrame {filename}: {e}")
                continue

        return corrected_data

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
                    corrected_series = metric_values.apply(lambda x: self.correct(str(x), relevant_metrics) if pd.notnull(x) else x)
                    
                    print(f"\nProcessing file: {filename}")
                    print("Original unique metrics:", unique_metrics)
                    print("Corrected unique metrics:", relevant_metrics)

                    if len(corrected_series) == len(df.iloc[:, 0]):
                        df.iloc[:, 0] = corrected_series
                    else:
                        print(f"Length mismatch in file {filename}")
                        continue

                    output_filepath = os.path.join(directory_output, filename)
                    os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
                    
                    with pd.ExcelWriter(output_filepath, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, header=False)

                except (IndexError, ValueError) as e:
                    print(f"Problem with file {filename}: {e}")
                    problem_file_path = os.path.join(problem_directory, filename)
                    os.makedirs(problem_directory, exist_ok=True)
                    
                    if os.path.exists(problem_file_path):
                        os.remove(problem_file_path)
                    os.rename(filepath, problem_file_path)