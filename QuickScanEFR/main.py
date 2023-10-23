import os
from PdfToImageConverter_optimized import convert_pdfs_to_images
from ImagePreprocessor_optimized import process_images_in_folder
from OCRProcessor_optimized import process_texts_in_folder as ocr_process
import time

def delete_intermediate_files(folder_path):
    # Define the path for the converted images folder
    converted_images_folder = os.path.join(folder_path, "converted_images")
    
    # Traverse directory tree using os.walk()
    for dirpath, dirnames, filenames in os.walk(converted_images_folder, topdown=False):
        # Delete files which are not .xlsx
        for filename in filenames:
            if not filename.endswith('.xlsx'):
                os.remove(os.path.join(dirpath, filename))
        
        # Try to remove the directories if they are empty
        for dirname in dirnames:
            try:
                os.rmdir(os.path.join(dirpath, dirname))
            except:
                print(f"Le dossier {os.path.join(dirpath, dirname)} n'est pas vide, certains fichiers n'ont pas été supprimés.")
    
    # Try to remove the main converted_images folder if it's empty
    try:
        os.rmdir(converted_images_folder)
    except:
        print("Le dossier converted_images n'est pas vide, certains fichiers n'ont pas été supprimés.")

if __name__ == "__main__":
    start_time = time.time()  # Record the start time

    folder_path = input("Enter the path of the folder containing the PDFs: ")
    
    convert_pdfs_to_images(folder_path)
    print(f"PDFs converted in: {folder_path}/converted_images")

    process_images_in_folder(os.path.join(folder_path, "converted_images"))
    print("Images preprocessed.")

    ocr_process(os.path.join(folder_path, "converted_images", "preprocessed_images"))
    print("Text extracted and saved to Excel.")

    # Supprimer les fichiers intermédiaires
    delete_intermediate_files(folder_path)
    print("Intermediate files deleted.")

    end_time = time.time()  # Record the end time

    duration = end_time - start_time  # Calculate the duration
    print(f"The pipeline took {duration:.2f} seconds to complete.")