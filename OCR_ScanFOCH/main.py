import os 
from PdfToImageConverter import convert_pdfs_to_images
from ImagePreprocessor import process_images_in_folder
from OCRProcessor import process_texts_in_folder as ocr_process

if __name__ == "__main__":
    folder_path = input("Enter the path of the folder containing the PDFs: ")
    convert_pdfs_to_images(folder_path)
    print(f"PDFs converted in: {folder_path}/converted_images")

    process_images_in_folder(os.path.join(folder_path, "converted_images"))
    print("Images preprocessed.")

    ocr_process(os.path.join(folder_path, "converted_images", "preprocessed_images"))
    print("Text extracted and saved to Excel.")
