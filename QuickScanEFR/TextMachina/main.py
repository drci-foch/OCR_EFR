import time
from DateFormatter_refactored import DateFormatter
from PercentageExtractor_refactored import PercentageExtractor
from OutlierDetector_refactored import OutlierDetector
from refactored_VEMSImputer import VEMSCorrector
from refactored_UnitCorrector import UnitCorrector
from refactored_MetricTypo import TypoCorrector
from refactored_ExcelCleanup import ExcelCleanup
from PDFPlumber_refactored import PDFProcessor

class MainPipeline:
    def __init__(self, input_directory='../pdf_TextMachina', output_directory='../pdf_TextMachina/excel'):
        self.input_directory = input_directory
        self.output_directory = output_directory
        self.standardized_metrics = ['CV', 'CVF', 'VR Plethysmo', 'CPT Plethysmo', 'VEMS']

        # Initialize helper classes
        self.date_formatter = DateFormatter()
        self.percentage_extractor = PercentageExtractor()
        self.outlier_detector = OutlierDetector()
        self.vems_imputer = VEMSCorrector()
        self.unit_corrector = UnitCorrector()
        self.typo_corrector = TypoCorrector(self.standardized_metrics)
        self.excel_cleaner = ExcelCleanup()
        self.pdf_processor = PDFProcessor(directory_path=input_directory)

    def process_pdfs(self):
        self.pdf_processor.process_directory()
        #print("process_pdfs done")

    def format_dates(self):
        self.date_formatter.format_dates_in_excel(directory_path=self.output_directory)
        #print("format_dates done")
        
    def extract_percentages(self):
        self.percentage_extractor.extract_percentages_from_excel(directory_path=self.output_directory)
        #print("extract_percentages done")

    def detect_anomalies(self):
        self.outlier_detector.detect_anomalies_for_multiple_docs(directory_path=self.output_directory)
        #print("detect_anomalies done")

    def compute_missing_vems(self):
        self.vems_imputer.correct_multiple_docs(directory_path=self.output_directory)
        #print("compute_missing_vems done")

    def correct_units(self):
        self.unit_corrector.correct_multiple_docs(directory_path=self.output_directory)
        #print("correct_units done")

    def correct_metric_typos(self):
        self.typo_corrector.correct_multiple_docs(directory_path=self.output_directory)
        #print("correct_metric_typos done")

    def clean_excel_files(self):
        self.excel_cleaner.clean_multiple_docs(directory_path=self.output_directory)
        #print("clean_excel_files done")

    def run(self):
        self.process_pdfs()
        self.correct_metric_typos()
        self.format_dates()
        self.extract_percentages()
        self.detect_anomalies()
        self.compute_missing_vems()
        self.correct_units()
        self.clean_excel_files()

if __name__ == "__main__":
    start_time = time.time() 
    pipeline = MainPipeline()
    pipeline.run()
    end_time = time.time()  
    elapsed_time = end_time - start_time 
    print(f"Pipeline took {elapsed_time:.2f} seconds to complete.")