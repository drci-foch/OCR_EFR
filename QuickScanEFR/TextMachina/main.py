from PercentageExtractor import PercentageExtractor
from DateFromatter import DateFormatter
from PDFPlumber import PDFProcessor
from OutlierDetector import OutlierDetector


directory_path = "../pdf_TextMachina/"

# Initialization
date_formatter = DateFormatter()
percent_extractor = PercentageExtractor()
pdf_processor = PDFProcessor(directory_path=directory_path)
outlier_detector = OutlierDetector(None)

# Process the PDFs to extract tables and save them in Excel format
pdf_processor.process_directory()

# Format the dates
date_formatter.format_dates_in_excel(directory_path)

# Extract the percentages
percent_extractor.extract_percentages_from_excel(directory_path)

# Use the OutlierDetector
outlier_detector.detect_anomalies_for_multiple_docs(directory_path)
