import os
import cv2
import numpy as np
from concurrent.futures import ProcessPoolExecutor


class ImagePreprocessor:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.output_folder = os.path.join(folder_path)

    def _preprocess_image(self, image_path):
        image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
        image = cv2.resize(image, None, fx=4, fy=4,
                           interpolation=cv2.INTER_CUBIC)
        _, image = cv2.threshold(image, 100, 400, cv2.THRESH_BINARY)
        kernel = np.ones((5, 5), np.uint8)  # A 5x5 kernel of ones
        image = cv2.erode(image, kernel, iterations=1)
        image = image[1750*2:3400*2, 0:-1]

        base_name = os.path.splitext(os.path.basename(image_path))[0]
        preprocessed_image_path = os.path.join(
            self.output_folder, f'{base_name}.png')
        cv2.imwrite(preprocessed_image_path, image)
      

    def process_images_in_folder(self):
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)

        png_files = [file for file in os.listdir(
            self.folder_path) if file.lower().endswith('.png')]
        with ProcessPoolExecutor() as executor:
            results = list(executor.map(self._preprocess_image,
                                        [os.path.join(self.folder_path, png_file) for png_file in png_files]))

            print(results)


if __name__ == "__main__":
    pipeline = ImagePreprocessor()
    pipeline.process_images_in_folder()