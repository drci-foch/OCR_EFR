import os
import cv2
import pytesseract
import pandas as pd
from concurrent.futures import ProcessPoolExecutor

def extract_text_from_image(image_path):
    image = cv2.imread(image_path)
    config = ''  # Utilisez le mode LSTM de Tesseract
    return pytesseract.image_to_string(image, config=config)

def combine_occurrences(sections):
    i = 0
    while i < len(sections) - 1:
        pair = (sections[i], sections[i+1])
        if pair in COMBINE_PAIRS:
            sections[i] = " ".join(pair)
            sections.pop(i + 1)
        else:
            i += 1
    return sections

def extract_section(sections, start_keywords=None, end_keywords=None, end_offset=0):
    if start_keywords:
        start_idx = next((sections.index(keyword) + 1 for keyword in start_keywords if keyword in sections), 0)
    else:
        start_idx = 0
    if end_keywords:
        end_idx = next((sections.index(keyword) + end_offset for keyword in end_keywords if keyword in sections), len(sections))
    else:
        end_idx = len(sections)
    return [item for item in sections[start_idx:end_idx] if item]

def adjust_list_length(target_list, reference_list):
    target_list = ["", ""] + target_list
    while len(target_list) < len(reference_list):
        target_list.append("")
    return target_list

def process_file(image_path):
    text = extract_text_from_image(image_path)
    sections = text.split()
    
    sections_set = set(sections)
    end_keyword = ['VIMS'] if 'VIMS' in sections_set else ['DEM25']

    sections = combine_occurrences(sections)
    parameters_section = extract_section(sections, None, end_keyword, 1)        
    theo_values = extract_section(sections, THEO_KEYWORDS, PRE_KEYWORDS)
    pre_values = extract_section(sections, PRE_KEYWORDS, PERC_THEO_KEYWORDS)
    perc_theo_values = extract_section(sections, PERC_THEO_KEYWORDS)
    
    theo_values = adjust_list_length(theo_values, parameters_section)
    perc_theo_values = adjust_list_length(perc_theo_values, parameters_section)
    
    df_data = {
        'Paramètres': parameters_section,
        'Théo': theo_values,
        'Pré': pre_values,
        '% Théo': perc_theo_values
    }

    return df_data

def process_texts_in_folder(folder_path):
    files = os.listdir(folder_path)
    png_files = [file for file in files if file.lower().endswith('.png')]
    
    dfs = []
    with ProcessPoolExecutor() as executor:
        results = executor.map(process_file, [os.path.join(folder_path, png_file) for png_file in png_files])

        for png_file, df_data in zip(png_files, results):
            dfs.append(pd.DataFrame(df_data))
            base_name = os.path.splitext(png_file)[0]
            excel_filename = os.path.join(folder_path, f"{base_name}_excel.xlsx")
            dfs[-1].to_excel(excel_filename, index=False)
            print(f"Saved data from {png_file} to {excel_filename}")

COMBINE_PAIRS = {
    ("Date", "test"),
    ("Heure", "test"),
    ('%', 'Théo'),
    ('*', 'Théo'),
    ('%', 'Theo'),
    ('*', 'Theo'),
    ('‘*', 'Theo'),
    ('CV', 'Max')
}

THEO_KEYWORDS = set(['Théo', 'Theo', ' Théo', ' Theo', 'Théo ', 'Theo '])
PRE_KEYWORDS = set(['Pre', 'Pré', 'Pre ', 'Pré ', ' Pré', ' Pre'])
PERC_THEO_KEYWORDS = set(['%Théo','% Théo', "%Theo", "% Theo", '*Théo','* Théo', "*Theo", "* Theo", '‘* Theo'])

if __name__ == "__main__":
    folder_path = input("Enter the path of the folder containing the preprocessed images: ")
    process_texts_in_folder(folder_path)
