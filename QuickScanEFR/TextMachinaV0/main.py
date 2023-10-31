import os
import pandas as pd
from PercentageExtractor import PercentageExtractor
from DateFormatter import DateFormatter
from PDFPlumber import PDFProcessor
from OutlierDetector import OutlierDetector
from VEMSImputer import VEMSCorrector
from UnitCorrector import UnitCorrector
from ExcelCleanup import ExcelCleanup
from MetricTypo import TypoCorrector

standardized_metrics = [
    "CV", "CVF", "VR Helium", "VR  Plethysmo", "CPT Helium", "CPT Plethysmo",
    "VR/CPT", "VEMS", "VEMS/CV", "DEM 25-75", "AA/O²", "PaO²", "PaCO²", "pH",
    "TLCO", "TLCO/Va", "AA/O²", "Walk Test", "SaO² mini", "PI", "Arrêt inter"
]

class MainPipeline:
    def __init__(self, directory_path):
        self.directory_path = directory_path
        self.date_formatter = DateFormatter()
        self.percent_extractor = PercentageExtractor()
        self.pdf_processor = PDFProcessor(directory_path=self.directory_path)
        self.outlier_detector = OutlierDetector(None)
        self.vems_corrector = VEMSCorrector()
        self.unit_corrector = UnitCorrector()
        self.excel_cleanup = ExcelCleanup()
        self.typo_corrector = TypoCorrector(standardized_metrics)


    def process_pdfs(self):
        """Extract tables from PDFs and save them in Excel format."""
        self.pdf_processor.process_directory()

    def format_dates(self):
        """Reformat the dates in the extracted Excel files."""
        self.date_formatter.format_dates_in_excel(self.directory_path)

    def extract_percentages(self):
        """Extract percentages from the Excel files."""
        self.percent_extractor.extract_percentages_from_excel(self.directory_path)

    def detect_anomalies(self):
        """Detect anomalies in the data."""
        self.outlier_detector.detect_anomalies_for_multiple_docs(self.directory_path)

    def correct_vems(self):
        """Compute missing VEMS values."""
        self.vems_corrector.correct_multiple_docs(self.directory_path)

    def correct_unit(self):
        """Correct the CVF values."""
        self.unit_corrector.correct_multiple_docs(self.directory_path)
        
    def clean_excel_files(self):
        self.excel_cleanup.clean_multiple_docs(self.directory_path)

    def correct_typos_in_excel(self):
        """Correct typos in the metric names in Excel files."""
        for filename in os.listdir(self.directory_path):
            if filename.endswith("_cleaned.xlsx"):
                filepath = os.path.join(self.directory_path, filename)
                df = pd.read_excel(filepath)
                
                # Extract unique metric names from the document
                unique_metrics = df[0].dropna().unique().tolist()
                
                # Determine the closest standardized names for each unique metric
                relevant_metrics = [self.typo_corrector.correct(metric, self.typo_corrector.standardized_strings) 
                                    for metric in unique_metrics]
                
                # Correct potential typos using the relevant metrics
                df[0] = df[0].apply(lambda x: self.typo_corrector.correct(str(x), relevant_metrics) if pd.notnull(x) else x)
                
                # Save the corrected dataframe back to the file using ExcelWriter
                with pd.ExcelWriter(filepath) as writer:
                    df.to_excel(writer, index=False)

                
    def cleanup(self):
        """Remove intermediate files after the entire pipeline execution."""
        intermediate_suffixes = ["_cleaned.xlsx", "_cleaned_modified.xlsx", "_corrected.xlsx", "_cvf.xlsx"]
        for filename in os.listdir(self.directory_path):
            for suffix in intermediate_suffixes:
                if filename.endswith(suffix):
                    os.remove(os.path.join(self.directory_path, filename))

    def run(self):
        self.process_pdfs()
        self.format_dates()
        self.correct_typos_in_excel()
        self.extract_percentages()
        self.detect_anomalies()
        self.correct_vems()
        self.correct_unit()
        self.clean_excel_files()
        self.cleanup()

# Usage:
pipeline = MainPipeline(directory_path="../pdf_TextMachina/")
pipeline.run()
