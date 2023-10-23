import os
import cv2
import numpy as np

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

def process_images_in_folder(folder_path):
    output_folder = os.path.join(folder_path, "preprocessed_images")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # List all files in the directory
    files = os.listdir(folder_path)

    # Filter out only PNG files
    png_files = [file for file in files if file.lower().endswith('.png')]

    for png_file in png_files:
        image_path = os.path.join(folder_path, png_file)
        preprocess_image(image_path, output_folder)
        print(f"Processed {png_file}")
