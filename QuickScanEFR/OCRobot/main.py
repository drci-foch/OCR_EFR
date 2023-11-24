import os
import time
import pandas as pd
from PdfToImageConverter_optimized import PdfConverter
from ImagePreprocessor_optimized import ImagePreprocessor
from OCRProcessor_optimized import TextExtractorFromImages
from ExcelFormatter import DataReshaper


class MainPipeline:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.pdf_converter = PdfConverter(folder_path)
        self.image_preprocessor = ImagePreprocessor(os.path.join(folder_path))
        self.text_extractor = TextExtractorFromImages(
            os.path.join(folder_path))

    def delete_intermediate_files(self):
        for dirpath, _, filenames in os.walk(self.folder_path, topdown=False):
            for filee in filenames:
                if filee.endswith('.png'):
                    os.remove(os.path.join(dirpath, filee))

    def reshape_data(self):
        generated_files_path = os.path.join(self.folder_path)

        for filename in os.listdir(generated_files_path):
            if filename.endswith('.xlsx'):
                file_path = os.path.join(generated_files_path, filename)
                df = pd.read_excel(file_path)
                reshaper = DataReshaper(df)
                reshaped_df = reshaper.reshape()
                reshaped_df.to_excel(file_path, index=False)

    def concatenate_excel_files(self):
        # List all files in the given directory
        all_files = os.listdir(self.folder_path)

        # Filter out the Excel files
        excel_files = [f for f in all_files if f.endswith('.xlsx')]

        # Group files by ID
        files_by_id = {}
        for file in excel_files:
            file_id = file.split('_')[0]
            if file_id not in files_by_id:
                files_by_id[file_id] = []
            files_by_id[file_id].append(file)

        # Create "concatenated" subdirectory if it doesn't exist
        concatenated_folder = os.path.join(self.folder_path, "concatenated")
        if not os.path.exists(concatenated_folder):
            os.makedirs(concatenated_folder)

        # Concatenate files with the same ID
        for file_id, files in files_by_id.items():
            dfs = []
            for file in files:
                dfs.append(pd.read_excel(os.path.join(self.folder_path, file)))

            concatenated_df = pd.concat(dfs, axis=0, ignore_index=True)

            # Save concatenated file
            concatenated_df.to_excel(os.path.join(
                concatenated_folder, f"{file_id}.xlsx"), index=False)

    def format_excel(self):
        concatenated_id = os.path.join(self.folder_path, "concatenated")
        all_files = os.listdir(concatenated_id)
        for filename in all_files:
            if filename.endswith(".xlsx"):
                file_path = os.path.join(concatenated_id, filename)

                df = pd.read_excel(file_path)

                df["Date"] = pd.to_datetime(
                    df["Date"], dayfirst=True).dt.strftime('%d/%m/%Y')

                # I assume that if % not in column name then it is in L
                no_perc = [
                    e for e in df.columns if "%" not in e and "Date" not in e]
                perc = [e for e in df.columns if "%" in e]
                for c in no_perc:
                    df[c] = df[c]*1000

                for c in perc:
                    df[c] = df[c]/100

                print(file_path)
                df.to_excel(file_path, index=False)

    def run(self):
        start_time = time.time()

        # Convert PDFs to Images
        self.pdf_converter.convert_pdfs_to_images()
        print(f"self.pdf_converter.convert_pdfs_to_images() done")

        # Preprocess Images
        self.image_preprocessor.process_images_in_folder()
        print(f"self.image_preprocessor.process_images_in_folder() done")

        # Extract Text from Images using OCR
        self.text_extractor.process_texts_in_folder()
        print(f"self.text_extractor.process_texts_in_folder() done")

        # Reshape the generated Excel data
        self.reshape_data()
        print(f"self.reshape_data() done")

        # Delete intermediate files
        self.delete_intermediate_files()
        print(f"self.delete_intermediate_files() done")

        # Concatenate excel file for each patient
        self.concatenate_excel_files()
        print(f"self.concatenate_excel_files() done")

        self.format_excel()
        print(f"self.format_excel() done")

        end_time = time.time()
        duration = end_time - start_time
        print(f"The pipeline took {duration:.2f} seconds to complete.")


# Test
if __name__ == "__main__":
    folder_path = "../pdf_OCRobot"
    pipeline = MainPipeline(folder_path)
    pipeline.run()
