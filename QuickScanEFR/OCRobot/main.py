import os
from PdfToImageConverter_optimized import PdfConverter
from ImagePreprocessor_optimized import ImagePreprocessor
from OCRProcessor_optimized import TextExtractorFromImages
import time


class MainPipeline:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.pdf_converter = PdfConverter(folder_path)
        self.image_preprocessor = ImagePreprocessor(
            os.path.join(folder_path, "converted_images"))
        self.text_extractor = TextExtractorFromImages(os.path.join(
            folder_path, "converted_images", "preprocessed_images"))

    def delete_intermediate_files(self):
        converted_images_folder = os.path.join(
            self.folder_path, "converted_images")

        for dirpath, dirnames, filenames in os.walk(converted_images_folder, topdown=False):
            for filename in filenames:
                if not filename.endswith('.xlsx'):
                    os.remove(os.path.join(dirpath, filename))

            for dirname in dirnames:
                try:
                    os.rmdir(os.path.join(dirpath, dirname))
                except:
                    print(
                        f"Le dossier {os.path.join(dirpath, dirname)} n'est pas vide, certains fichiers n'ont pas été supprimés.")

        try:
            os.rmdir(converted_images_folder)
        except:
            print(
                "Le dossier converted_images n'est pas vide, certains fichiers n'ont pas été supprimés.")

    def run(self):
        start_time = time.time()  # Record the start time

        # Convert PDFs to Images
        self.pdf_converter.convert_pdfs_to_images()
        print(f"PDFs converted in: {self.folder_path}/converted_images")

        # Preprocess Images
        self.image_preprocessor.process_images_in_folder()
        print("Images preprocessed.")

        # Extract Text from Images using OCR
        self.text_extractor.process_texts_in_folder()
        print("Text extracted and saved to Excel.")

        # Delete intermediate files
        self.delete_intermediate_files()
        print("Intermediate files deleted.")

        end_time = time.time()  # Record the end time
        duration = end_time - start_time  # Calculate the duration
        print(f"The pipeline took {duration:.2f} seconds to complete.")


# Test
if __name__ == "__main__":
    # Replace with your actual path
    folder_path = "/Users/sarrabenyahia/Documents/GitHub/OCR_EFR/test_OCRobot"
    pipeline = MainPipeline(folder_path)
    pipeline.run()
