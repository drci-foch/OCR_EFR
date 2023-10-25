from PercentageExtractor import PercentageExtractor
from DateFromatter import DateFormatter
from PDFPlumber import PDFProcessor 
from OutlierDetector import OutlierDetector

# Initialization
date_formatter = DateFormatter()
percent_extractor = PercentageExtractor()
pdf_processor = PDFProcessor(directory_path='../pdf_TextMachina/')

# Process the PDFs to extract tables and save them in Excel format
pdf_processor.process_directory()

# Format the dates
excel_file_path = "../pdf_TextMachina/"  # Replace with the actual Excel file name if needed
date_formatter.format_dates_in_excel(excel_file_path)

# Extract the percentages
percent_extractor.extract_percentages_from_excel(excel_file_path)

# Initialize and use the OutlierDetector for multiple documents
outlier_detector = OutlierDetector(None)
directory_path = "../pdf_TextMachina/"
outlier_detector.detect_anomalies_for_multiple_docs(directory_path)
