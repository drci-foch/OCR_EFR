import pytesseract
import cv2
import logging
import os


def ocr_image(image_path: str) -> str:
    """
    Extract text from an image using OCR.
    
    Args:
    - image_path (str): Path to the input image.
    
    Returns:
    - Extracted text from the image.
    """
    try:
        image = cv2.imread(image_path)
        text = pytesseract.image_to_string(image)
        return text
    except Exception as e:
        logging.error(f"Failed to OCR {image_path}. Error: {e}")
        return ""

def ocr_images_in_folder(folder_path: str, output_folder: str) -> None:
    """
    OCR all images in the specified folder and save the extracted text.
    
    Args:
    - folder_path (str): Path to the folder containing images.
    - output_folder (str): Folder to save the extracted text.
    """
    # List all files in the folder
    files = os.listdir(folder_path)

    # Filter out image files
    image_files = [file for file in files if file.endswith(('.png', '.jpg', '.jpeg'))]

    for image_file in image_files:
        image_path = os.path.join(folder_path, image_file)
        extracted_text = ocr_image(image_path)
        
        # Save the extracted text to a file
        output_file_path = os.path.join(output_folder, f"{image_file[:-4]}.txt")
        with open(output_file_path, 'w', encoding='utf-8') as f:
            f.write(extracted_text)

#To execute
image_folder = "path_to_image_folder"
text_output_folder = "path_to_text_output_folder"
ocr_images_in_folder(image_folder, text_output_folder)
