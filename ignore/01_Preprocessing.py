import os
import cv2
import logging
from pdf2image import convert_from_path
#TODO : Change pdf_to_images to take into account the localisation of the image (do we want the patient's id ?)

# Initialize logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def pdf_to_images(pdf_path: str, output_folder: str) -> list:
    """
    Convert a PDF file to a set of images and return the paths to the generated images.
    
    Args:
    - pdf_path (str): Path to the input PDF file.
    - output_folder (str): Folder to save the output images.
    
    Returns:
    - List of paths to the generated images.
    """
    generated_images = []
    try:
        images = convert_from_path(pdf_path)
        for i, image in enumerate(images):
            image_path = os.path.join(output_folder, f'page_{i + 1}.png')
            image.save(image_path, 'PNG')
            generated_images.append(image_path)
    except Exception as e:
        logging.error(f"Failed to convert {pdf_path}. Error: {e}")
    return generated_images

def process_image(image_path: str, output_path: str) -> None:
    """
    Process the image and save the result.
    
    Args:
    - image_path (str): Path to the input image.
    - output_path (str): Path to save the processed image.
    """
    # Grayscale
    image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)

    # Binarisation
    _, image = cv2.threshold(image, 170, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Improve contrast with equalizeHist
    image = cv2.equalizeHist(image)

    cv2.imwrite(output_path, image)

def process_pdfs_in_folder(pdf_folder: str, image_output_folder: str) -> None:
    """
    Convert PDFs in a folder to images and process those images.
    
    Args:
    - pdf_folder (str): Path to the folder containing PDF files.
    - image_output_folder (str): Folder to save the output images from PDFs.
    """
    # List all files in the folder
    files = os.listdir(pdf_folder)

    # Filter out PDF files
    pdf_files = [file for file in files if file.endswith('.pdf')]

    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        generated_images = pdf_to_images(pdf_path, image_output_folder)
        
        # Process only the newly generated images from the current PDF
        for image_path in generated_images:
            output_path = os.path.join(image_output_folder, f"{os.path.basename(image_path)[:-4]}_OCR.png")
            if not os.path.exists(output_path):
                process_image(image_path, output_path)


# To execute (change paths)
pdf_folder = "path_to_pdf_folder"
image_output_folder = "path_to_image_output_folder"
process_pdfs_in_folder(pdf_folder, image_output_folder)
