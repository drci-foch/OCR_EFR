import os
from PercentageExtractor import PercentageExtractor
from DateFormatter import DateFormatter
from PDFPlumber import PDFProcessor
from OutlierDetector import OutlierDetector
from VEMSImputer import VEMSCorrector
from UnitCorrector import UnitCorrector
from ExcelCleanup import ExcelCleanup

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
        self.extract_percentages()
        self.detect_anomalies()
        self.correct_vems()
        self.correct_unit()
        self.clean_excel_files()
        self.cleanup()

# Usage:
pipeline = MainPipeline(directory_path="../pdf_TextMachina/")
pipeline.run()
