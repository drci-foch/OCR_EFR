import os
from pdf2image import convert_from_path

def convert_pdfs_to_images(folder_path):
    output_folder = os.path.join(folder_path, "converted_images")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # List all files in the directory
    files = os.listdir(folder_path)

    # Filter out only PDF files
    pdf_files = [file for file in files if file.lower().endswith('.pdf')]

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        # Removing the '.pdf' extension and adding a prefix for image names
        base_name = os.path.splitext(pdf_file)[0] + '_'

        # Convert only the first page of the PDF
        images = convert_from_path(pdf_path, first_page=1, last_page=1)
        for i, image in enumerate(images):
            image_path = os.path.join(output_folder, f'{base_name}{i}.png')
            image.save(image_path, 'PNG')
        print(f"Processed {pdf_file}")

    return folder_path
