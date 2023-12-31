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
    def __init__(self, input_directory='C:/Users/benysar/Desktop/Github/OCR_EFR/QuickScanEFR/pdf_TextMachina/excel_correctedVEMS/', output_directory='C:/Users/benysar/Desktop/Github/OCR_EFR/QuickScanEFR/pdf_TextMachina/excel_correctedVEMS'):
        
        self.input_directory = input_directory
        self.output_directory = output_directory
        self.standardized_metrics = [
            "CV", "CVF", "VR Helium", "VR  Plethysmo", "CPT Helium", "CPT Plethysmo","Capacité Inspiratoire",
            "VR/CPT", "VEMS", "VEMS/CV", "DEM 25-75", "DEM 75", "DEM 25", "DEM 55", "Indice de dyspnée",
            "AA/O²", "PaO²", "PaCO²", "pH", "Sat. Mini.", "FC", "Sat", "Traitement", "SAO² initiale",
            "TLCO", "TLCO/Va", "DLCO", "DLCO/Va", "AA/O²", "Walk Test", "SaO² mini", "PI", "Arrêt inter", "Arrêt",
            "FC Max","date","HCO3-","Saturation", "Walk test \n Sat min-max \n Fc min-max", "CVL", "R5Hz", 
            "VEMS après Vent", "Echelle dyspnée",

        ]

        # Initialize helper classes
        self.date_formatter = DateFormatter()
        self.percentage_extractor = PercentageExtractor()
        self.outlier_detector = OutlierDetector()
        self.vems_imputer = VEMSCorrector()
        self.unit_corrector = UnitCorrector()
        self.typo_corrector = TypoCorrector(self.standardized_metrics)
        self.excel_cleaner = ExcelCleanup()
        self.pdf_processor = PDFProcessor(directory_path=input_directory, output_path=output_directory)

    def process_pdfs(self):
        start_time = time.time()
        self.pdf_processor.process_directory()
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"process_pdfs took {elapsed_time:.2f} seconds")

    def format_dates(self):
        start_time = time.time()
        self.date_formatter.format_dates_in_excel(
            directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"format_dates took {elapsed_time:.2f} seconds")

    def extract_percentages(self):
        start_time = time.time()
        self.percentage_extractor.extract_percentages_from_excel(
            directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"extract_percentages took {elapsed_time:.2f} seconds")

    def detect_anomalies(self):
        start_time = time.time()
        self.outlier_detector.detect_anomalies_for_multiple_docs(
            directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"detect_anomalies took {elapsed_time:.2f} seconds")

    def compute_missing_vems(self):
        start_time = time.time()
        self.vems_imputer.correct_multiple_docs(
            directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"compute_missing_vems took {elapsed_time:.2f} seconds")

    def correct_units(self):
        start_time = time.time()
        self.unit_corrector.correct_multiple_docs(
            directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"correct_units took {elapsed_time:.2f} seconds")

    def correct_metric_typos(self):
        start_time = time.time()
        self.typo_corrector.correct_multiple_docs(
            directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"correct_metric_typos took {elapsed_time:.2f} seconds")

    def clean_excel_files(self):
        start_time = time.time()
        self.excel_cleaner.clean_multiple_docs(
            directory_path=self.output_directory)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"clean_excel_files took {elapsed_time:.2f} seconds")

    def run(self):
        # self.process_pdfs()
        # self.correct_metric_typos()
        # self.format_dates()
        # self.extract_percentages()
        # self.correct_units()
        # self.compute_missing_vems()
        self.clean_excel_files()
        # self.detect_anomalies()


if __name__ == "__main__":
    start_time = time.time()
    pipeline = MainPipeline()
    pipeline.run()
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Pipeline took {elapsed_time:.2f} seconds to complete.")
