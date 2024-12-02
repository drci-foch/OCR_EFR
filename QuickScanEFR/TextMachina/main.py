import time
import os 
import shutil
from DateFormatter_refactored import DateFormatter
from PercentageExtractor_refactored import PercentageExtractor
from OutlierDetector_refactored import OutlierDetector
from refactored_VEMSImputer import VEMSCorrector
from refactored_UnitCorrector import UnitCorrector
from refactored_MetricTypo import TypoCorrector
from refactored_ExcelCleanup import ExcelCleanup
from PDFPlumber_refactored import PDFProcessor


class MainPipeline:
    def __init__(self, input_directory=r'C:\Users\benysar\Documents\GitHub\OCR_EFR\QuickScanEFR\TextMachina\pdf', output_directory=r'C:\Users\benysar\Documents\GitHub\OCR_EFR\QuickScanEFR\TextMachina\pdf_test'):
        
        self.input_directory = input_directory
        self.output_directory = output_directory

        self.pdf_output = os.path.join(self.output_directory, 'pdf_extracted')
        self.typo_output = os.path.join(self.output_directory, 'typo_corrected')
        self.dates_output = os.path.join(self.output_directory, 'dates_formatted')
        self.percentage_output = os.path.join(self.output_directory, 'percentage_extracted')
        self.correct_output = os.path.join(self.output_directory, 'correct_units')
        self.clean_output = os.path.join(self.output_directory, 'final_clean_excel')
        # Create all directories
        for directory in [self.pdf_output, self.typo_output, self.dates_output, self.percentage_output, self.correct_output, self.output_directory]:
            os.makedirs(directory, exist_ok=True)

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
        # Process PDFs first
        self.pdf_processor.process_directory()
        
        # Now copy the processed files to the pdf_extracted folder
        for filename in os.listdir(self.output_directory):
            if filename.endswith('.xlsx'):
                source_path = os.path.join(self.output_directory, filename)
                dest_path = os.path.join(self.pdf_output, filename)
                
                # Create output directory if it doesn't exist
                os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                
                # Copy the file
                shutil.copy2(source_path, dest_path)
                print(f"Saved processed PDF data to: {dest_path}")
                
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"process_pdfs took {elapsed_time:.2f} seconds")

    def correct_metric_typos(self):
        """Correct typos in files from pdf_output and save to typo_output"""
        start_time = time.time()
        self.typo_corrector.correct_multiple_docs(
            directory_path=self.pdf_output,  # Read from previous step
            directory_output=self.typo_output  # Save to new directory
        )
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"correct_metric_typos took {elapsed_time:.2f} seconds")

    def format_dates(self):
        """Format dates in files from typo_output and save to dates_output"""
        start_time = time.time()
        self.date_formatter.format_dates_in_excel(
            directory_path=self.typo_output,  # Read from previous step
            output_path=self.dates_output  # Save to new directory
        )
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"format_dates took {elapsed_time:.2f} seconds")

    def extract_percentages(self):
        start_time = time.time()
        self.percentage_extractor.extract_percentages_from_excel(
            directory_path=self.dates_output,
            output_path=self.percentage_output )
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"extract_percentages took {elapsed_time:.2f} seconds")

    def correct_units(self):
        start_time = time.time()
        self.unit_corrector.correct_multiple_docs(
            directory_path=self.percentage_output,
            output_path=self.correct_output)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"correct_units took {elapsed_time:.2f} seconds")

    # def compute_missing_vems(self):
    #     start_time = time.time()
    #     self.vems_imputer.correct_multiple_docs(
    #         directory_path=self.output_directory)
    #     end_time = time.time()
    #     elapsed_time = end_time - start_time
    #     print(f"compute_missing_vems took {elapsed_time:.2f} seconds")
    
    # def detect_anomalies(self):
    #     start_time = time.time()
    #     self.outlier_detector.detect_anomalies_for_multiple_docs(
    #         directory_path=self.output_directory)
    #     end_time = time.time()
    #     elapsed_time = end_time - start_time
    #     print(f"detect_anomalies took {elapsed_time:.2f} seconds")

    def clean_excel_files(self):
        start_time = time.time()
        self.excel_cleaner.clean_multiple_docs(
            directory_path=self.correct_output,  # Read from previous step
            output_directory=self.clean_output   # Save to final clean excel directory
        )
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"clean_excel_files took {elapsed_time:.2f} seconds")

    def run(self):
        """Run the complete pipeline sequentially"""
        self.process_pdfs() 
        self.correct_metric_typos() 
        self.format_dates() 
        self.extract_percentages()
        self.correct_units()
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
