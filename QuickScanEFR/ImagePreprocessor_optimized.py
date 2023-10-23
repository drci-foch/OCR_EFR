import os
import cv2
import numpy as np
from concurrent.futures import ProcessPoolExecutor

def preprocess_image(image_path, output_folder):
    # Lire l'image en niveau de gris
    image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)

    # Redimensionner l'image
    image = cv2.resize(image, None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)

    # Binarisation avec un seuil fixe
    _, image = cv2.threshold(image, 100, 400, cv2.THRESH_BINARY)

    # Define a kernel for erosion. This can be adjusted based on your needs.
    kernel = np.ones((5,5),np.uint8)  # A 5x5 kernel of ones

    # Apply erosion
    image = cv2.erode(image, kernel, iterations = 1)
    
    # Crop l'image sur le tableau d'intérêt
    image = image[1750*2:3400*2, 0:-1]

    # Enregistrement de l'image prétraitée
    base_name = os.path.splitext(os.path.basename(image_path))[0]
    preprocessed_image_path = os.path.join(output_folder, f'{base_name}_preprocessed.png')
    cv2.imwrite(preprocessed_image_path, image)
    return f"Processed {base_name}"

def process_images_in_folder(folder_path):
    output_folder = os.path.join(folder_path, "preprocessed_images")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # List and filter PNG files in the directory in one operation
    png_files = [file for file in os.listdir(folder_path) if file.lower().endswith('.png')]

    with ProcessPoolExecutor() as executor:
        results = list(executor.map(preprocess_image, 
                                    [os.path.join(folder_path, png_file) for png_file in png_files], 
                                    [output_folder]*len(png_files)))
        
        for message in results:
            print(message)

if __name__ == "__main__":
    folder_path = input("Enter the path of the folder containing the images to preprocess: ")
    process_images_in_folder(folder_path)
