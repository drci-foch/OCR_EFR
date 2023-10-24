from PercentageExtractor import PercentageExtractor
from DateFromatter import DateFormatter
from PDFPlumber import PDFProcessor  # Assuming the class is in PDFProcessor.py

# Initialization
date_formatter = DateFormatter()
percent_extractor = PercentageExtractor()
pdf_processor = PDFProcessor(directory_path='./pdf')

# First, process the PDFs to extract tables and save them in Excel format
pdf_processor.process_directory()

# Then, format the dates
date_formatter.format_dates_in_excel("./pdf/")

# Finally, extract the percentages
percent_extractor.extract_percentages_from_excel("./pdf/")
