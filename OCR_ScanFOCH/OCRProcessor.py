import os
import cv2
import pytesseract
import pandas as pd

def extract_text_from_image(image_path):
    image = cv2.imread(image_path)
    
    # Utilisez le mode LSTM de Tesseract
    config = ''
    
    return pytesseract.image_to_string(image, config=config)


def combine_all_occurrences(sections, primary, secondary):
    """Combine all occurrences of two elements into a single string in the sections list."""
    i = 0
    while i < len(sections) - 1:  # -1 because we're checking pairs of elements
        if sections[i] == primary and sections[i + 1] == secondary:
            sections[i] = primary + " " + secondary
            sections.pop(i + 1)
        else:
            i += 1
    return sections

def extract_section(sections, start_keywords=None, end_keywords=None, end_offset=0):
    """
    Extract a section from the sections list based on start and end keywords.

    Args:
    - sections (list): List of text sections.
    - start_keywords (list, optional): Keywords marking the start of the section. 
                                       If None, start from the beginning.
    - end_keywords (list, optional): Keywords marking the end of the section.
                                     If None, end at the last item.
    - end_offset (int, optional): Offset to adjust the end index.

    Returns:
    - List of items in the extracted section.
    """
    
    # Combine elements
    sections = combine_all_occurrences(sections, "Date", "test")
    sections = combine_all_occurrences(sections, "Heure", "test")
    sections = combine_all_occurrences(sections, '%', 'Théo')
    sections = combine_all_occurrences(sections, '*', 'Théo')
    sections = combine_all_occurrences(sections, '%', 'Theo')
    sections = combine_all_occurrences(sections, '*', 'Theo')
    sections = combine_all_occurrences(sections, '‘*', 'Theo')
    sections = combine_all_occurrences(sections, 'CV', 'Max')

    
    if start_keywords:
        # Trouvez l'index du premier mot-clé correspondant
        start_idx = next((sections.index(
            keyword) + 1 for keyword in start_keywords if keyword in sections), 0)
    else:
        start_idx = 0

    if end_keywords:
        # Trouvez l'index du premier mot-clé correspondant
        end_idx = next((sections.index(
            keyword) + end_offset for keyword in end_keywords if keyword in sections), len(sections))
    else:
        end_idx = len(sections)
    return [item for item in sections[start_idx:end_idx] if item]


def adjust_list_length(target_list, reference_list):
    """Insert empty strings at the beginning of target_list for "Date test" and "Heure test"."""

    # Insert two empty strings at the beginning for "Date test" and "Heure test"
    target_list = ["", ""] + target_list

    # Ensure theo_values has at least as many values as reference_list
    while len(target_list) < len(reference_list):
        target_list.append("")

    return target_list


def process_texts_in_folder(folder_path):
    # List all files in the directory
    files = os.listdir(folder_path)

    # Filter out only PNG files
    png_files = [file for file in files if file.lower().endswith('.png')]

    for png_file in png_files:
        image_path = os.path.join(folder_path, png_file)
        text = extract_text_from_image(image_path)

        sections = text.split()  # You may need to adjust the split method if sections are separated differently.
        print(sections)
        end_keyword = ['VIMS'] if 'VIMS' in sections else ['DEM25']
        parameters_section = extract_section(sections, None, end_keyword, 1)        
        theo_values = extract_section(sections, ['Théo', 'Theo', ' Théo', ' Theo', 'Théo ', 'Theo '], ['Pre', 'Pré', 'Pre ', 'Pré ', ' Pré', ' Pre'])
        pre_values = extract_section(sections, ['Pre', 'Pré', 'Pre ', 'Pré ', ' Pré', ' Pre'], ['%Théo','% Théo', "%Theo", "% Theo", '*Théo','* Théo', "*Theo", "* Theo", '‘* Theo'])
        perc_theo_values = extract_section(sections, ['%Théo','% Théo', "%Theo", "% Theo", '*Théo','* Théo', "*Theo", "* Theo", '‘* Theo'])

        theo_values = adjust_list_length(theo_values, parameters_section)
        perc_theo_values = adjust_list_length(perc_theo_values, parameters_section)
        
        print(perc_theo_values)
        
        df_data = {
            'Paramètres': parameters_section,
            'Théo': theo_values,
            'Pré': pre_values,
            '% Théo': perc_theo_values
        }

        df = pd.DataFrame(df_data)

        # Saving the dataframe to an Excel file
        base_name = os.path.splitext(png_file)[0]
        excel_filename = os.path.join(folder_path, f"{base_name}_excel.xlsx")
        df.to_excel(excel_filename, index=False)

        print(f"Saved data from {png_file} to {excel_filename}")


if __name__ == "__main__":
    folder_path = input("Enter the path of the folder containing the preprocessed images: ")
    process_texts_in_folder(folder_path)
