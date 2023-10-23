import os
from pdf2image import convert_from_path
from concurrent.futures import ProcessPoolExecutor

def convert_pdf_to_image(pdf_path, output_folder):
    # Removing the '.pdf' extension and adding a prefix for image names
    base_name = os.path.splitext(os.path.basename(pdf_path))[0] + '_'

    # Convert only the first page of the PDF
    images = convert_from_path(pdf_path, first_page=1, last_page=1)
    for i, image in enumerate(images):
        image_path = os.path.join(output_folder, f'{base_name}{i}.png')
        image.save(image_path, 'PNG')
    return f"Processed {os.path.basename(pdf_path)}"

def convert_pdfs_to_images(folder_path):
    output_folder = os.path.join(folder_path, "converted_images")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # List and filter PDF files in the directory in one operation
    pdf_files = [file for file in os.listdir(folder_path) if file.lower().endswith('.pdf')]

    with ProcessPoolExecutor() as executor:
        results = list(executor.map(convert_pdf_to_image, 
                                    [os.path.join(folder_path, pdf_file) for pdf_file in pdf_files], 
                                    [output_folder]*len(pdf_files)))
        
        for message in results:
            print(message)

    return folder_path

if __name__ == "__main__":
    folder_path = input("Enter the path of the folder containing the PDFs to convert: ")
    convert_pdfs_to_images(folder_path)
