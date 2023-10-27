from excel_utils import load_excel_workbook, save_excel_data

# Integrate all the refactored classes
from PDFPlumber import PDFProcessorRefactored
from ExcelCleanup import ExcelCleanupRefactored
from PercentageExtractor import PercentageExtractorRefactored
from DateFormatter import DateFormatterRefactored
from OutlierDetector import OutlierDetectorRefactored
from CVFCorrector import CVFCorrectorRefactored
from VEMSImputer import VEMSCorrectorRefactored 

def main_processing_pipeline(directory_path):
    # Step 1: PDF to Excel Conversion
    pdf_processor = PDFProcessorRefactored()
    pdf_processor.process_directory(directory_path)
    
    # Step 2: Clean Excel Data
    excel_cleaner = ExcelCleanupRefactored()
    excel_cleaner.clean_multiple_docs(directory_path)

    # Step 3: Extract Percentage Values
    percentage_extractor = PercentageExtractorRefactored()
    percentage_extractor.extract_percentages_from_excel(directory_path)
    
    # Step 4: Format Dates
    date_formatter = DateFormatterRefactored()
    date_formatter.format_dates_in_excel(directory_path)
    
    # Step 5: VEMS Imputation 
    vems_corrector = VEMSCorrectorRefactored()
    vems_corrector.correct_multiple_docs(directory_path)
    
    # Step 6: Detect Anomalies
    outlier_detector = OutlierDetectorRefactored()
    outlier_detector.detect_anomalies_for_multiple_docs(directory_path)
    
    # Step 7: Correct CVF Values
    cvf_corrector = CVFCorrectorRefactored()
    cvf_corrector.correct_multiple_docs(directory_path)

    print("Processing complete!")


main_processing_pipeline("../pdf_TextMachina/")
