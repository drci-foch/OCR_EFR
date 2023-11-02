import os
from pdf2image import convert_from_path
from concurrent.futures import ProcessPoolExecutor


class PdfConverter:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.output_folder = os.path.join(folder_path)

    def convert_pdf_to_image(self, pdf_path, output_folder):
        base_name = os.path.splitext(os.path.basename(pdf_path))[0] + '_'
        images = convert_from_path(pdf_path, first_page=1, last_page=1)
        for i, image in enumerate(images):
            image_path = os.path.join(output_folder, f'{base_name}{i}.png')
            image.save(image_path, 'PNG')

    def convert_pdfs_to_images(self):
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)

        pdf_files = [file for file in os.listdir(
            self.folder_path) if file.lower().endswith('.pdf')]
        with ProcessPoolExecutor() as executor:
            results = list(executor.map(self.convert_pdf_to_image,
                                        [os.path.join(self.folder_path, pdf_file)
                                         for pdf_file in pdf_files],
                                        [self.output_folder]*len(pdf_files)))


