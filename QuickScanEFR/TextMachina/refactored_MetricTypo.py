import os
import pandas as pd
from typing import List

class TypoCorrector:
    def __init__(self, standardized_strings: List[str]):
        self.standardized_strings = standardized_strings

    @staticmethod
    def levenshtein_distance(s1: str, s2: str) -> int:
        if len(s1) > len(s2):
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
        return best_match if best_distance <= max_distance else string
    
    def correct_multiple_docs(self, directory_path: str):
        """Correct typos in the metric names in Excel files."""
        for filename in os.listdir(directory_path):
            if filename.endswith(".xlsx"):
                filepath = os.path.join(directory_path, filename)
                df = pd.read_excel(filepath)

                # Extract unique metric names from the document
                unique_metrics = df.iloc[:, 0].dropna().unique().tolist()

                # Determine the closest standardized names for each unique metric
                relevant_metrics = [self.correct(metric, self.standardized_strings) for metric in unique_metrics]
                
                # Correct potential typos using the relevant metrics
                df.iloc[:, 0] = df.iloc[:, 0].apply(lambda x: self.correct(str(x), relevant_metrics) if pd.notnull(x) else x)
                
                # Save the corrected dataframe back to the file using ExcelWriter
                with pd.ExcelWriter(filepath) as writer:
                    df.to_excel(writer, index=False, header=False)
