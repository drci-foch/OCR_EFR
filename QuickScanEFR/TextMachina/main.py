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
        start_time = time.time()
        self.pdf_processor.process_directory()
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"process_pdfs took {elapsed_time:.2f} seconds")

    def format_dates(self):
        start_time = time.time()
        self.date_formatter.format_dates_in_excel(directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"format_dates took {elapsed_time:.2f} seconds")

    def extract_percentages(self):
        start_time = time.time()
        self.percentage_extractor.extract_percentages_from_excel(directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"extract_percentages took {elapsed_time:.2f} seconds")

    def detect_anomalies(self):
        start_time = time.time()
        self.outlier_detector.detect_anomalies_for_multiple_docs(directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"detect_anomalies took {elapsed_time:.2f} seconds")

    def compute_missing_vems(self):
        start_time = time.time()
        self.vems_imputer.correct_multiple_docs(directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"compute_missing_vems took {elapsed_time:.2f} seconds")

    def correct_units(self):
        start_time = time.time()
        self.unit_corrector.correct_multiple_docs(directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"correct_units took {elapsed_time:.2f} seconds")

    def correct_metric_typos(self):
        start_time = time.time()
        self.typo_corrector.correct_multiple_docs(directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"correct_metric_typos took {elapsed_time:.2f} seconds")

    def clean_excel_files(self):
        start_time = time.time()
        self.excel_cleaner.clean_multiple_docs(directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"clean_excel_files took {elapsed_time:.2f} seconds")

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