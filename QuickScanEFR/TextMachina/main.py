from PercentageExtractor import PercentageExtractor
from DateFromatter import DateFormatter
from PDFPlumber import PDFProcessor 

# Initialization
date_formatter = DateFormatter()
percent_extractor = PercentageExtractor()
pdf_processor = PDFProcessor(directory_path='../QuickScanEFR/pdf_TextMachina')

# First, process the PDFs to extract tables and save them in Excel format
pdf_processor.process_directory()

# Then, format the dates
date_formatter.format_dates_in_excel("../QuickScanEFR/pdf_TextMachina")

# Finally, extract the percentages
percent_extractor.extract_percentages_from_excel("../QuickScanEFR/pdf_TextMachina")
