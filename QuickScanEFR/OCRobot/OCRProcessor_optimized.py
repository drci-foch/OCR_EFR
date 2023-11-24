import os
import cv2
import pytesseract
import pandas as pd
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\benysar\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'



class TextExtractorFromImages:

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
    PERC_THEO_KEYWORDS = set(
        ['%Théo', '% Théo', "%Theo", "% Theo", '*Théo', '* Théo', "*Theo", "* Theo", '‘* Theo'])

    def __init__(self, folder_path):
        self.folder_path = folder_path

    def _extract_text_from_image(self, image_path):
        image = cv2.imread(image_path)
        return pytesseract.image_to_string(image)

    def _combine_occurrences(self, sections):
        i = 0
        while i < len(sections) - 1:
            pair = (sections[i], sections[i+1])
            if pair in self.COMBINE_PAIRS:
                sections[i] = " ".join(pair)
                sections.pop(i + 1)
            else:
                i += 1
        return sections

    def _extract_section(self, sections, start_keywords=None, end_keywords=None, end_offset=0):
        if start_keywords:
            start_idx = next((sections.index(
                keyword) + 1 for keyword in start_keywords if keyword in sections), 0)
        else:
            start_idx = 0
        if end_keywords:
            end_idx = next((sections.index(
                keyword) + end_offset for keyword in end_keywords if keyword in sections), len(sections))
        else:
            end_idx = len(sections)
        return [item for item in sections[start_idx:end_idx] if item]

    def _adjust_list_length(self, target_list, reference_list):
        target_list = ["", ""] + target_list
        while len(target_list) < len(reference_list):
            target_list.append("")
        return target_list

    def _process_file(self, image_path):
        text = self._extract_text_from_image(image_path)
        sections = text.split()

        sections_set = set(sections)
        end_keyword = ['VIMS'] if 'VIMS' in sections_set else ['DEM25']

        sections = self._combine_occurrences(sections)
        parameters_section = self._extract_section(
            sections, None, end_keyword, 1)
        theo_values = self._extract_section(
            sections, self.THEO_KEYWORDS, self.PRE_KEYWORDS)
        pre_values = self._extract_section(
            sections, self.PRE_KEYWORDS, self.PERC_THEO_KEYWORDS)
        perc_theo_values = self._extract_section(
            sections, self.PERC_THEO_KEYWORDS)

        theo_values = self._adjust_list_length(theo_values, parameters_section)
        perc_theo_values = self._adjust_list_length(
            perc_theo_values, parameters_section)

        df_data = {
            'Paramètres': parameters_section,
            'Théo': theo_values,
            'Pré': pre_values,
            '% Théo': perc_theo_values
        }

        return df_data

    def process_texts_in_folder(self):
        files = os.listdir(self.folder_path)
        png_files = [file for file in files if file.lower().endswith('.png')]

        dfs = []
        for png_file in png_files:
            file_path = os.path.join(self.folder_path, png_file)
            df_data = self._process_file(file_path)
            
            df = pd.DataFrame(df_data)
            dfs.append(df)
            
            base_name = os.path.splitext(png_file)[0]
            excel_filename = os.path.join(self.folder_path, f"{base_name}.xlsx")
            df.to_excel(excel_filename, index=False)


if __name__ == "__main__":
    pipeline = TextExtractorFromImages()
    pipeline.process_texts_in_folder()
