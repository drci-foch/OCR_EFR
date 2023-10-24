# Auteur : Sarra Ben Yahia
# Point d'entrée principal de QuickScanEFR. Ce script vérifie les différents formats des fichiers PDF. Si le contenu du PDF est sélectionnable (c'est-à-dire rédigé manuellement par le médecin), alors TextMachina est exécuté. Dans le cas contraire, OCRobot est sollicité.

import os
import PyPDF2
import subprocess

# Directory paths
base_dir = os.path.dirname(os.path.abspath(__file__))
pdf_dir = os.path.join(base_dir, 'pdf')
pdf_textmachina_dir = os.path.join(base_dir, 'pdf_TextMachina')
pdf_ocrobot_dir = os.path.join(base_dir, 'pdf_OCRobot')

# Create the folders if they don't exist
if not os.path.exists(pdf_textmachina_dir):
    os.makedirs(pdf_textmachina_dir)

if not os.path.exists(pdf_ocrobot_dir):
    os.makedirs(pdf_ocrobot_dir)

# Check each PDF in the pdf directory
for filename in os.listdir(pdf_dir):
    if filename.endswith('.pdf'):
        filepath = os.path.join(pdf_dir, filename)
        
        with open(filepath, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            # Extract text from all pages of the PDF
            text = ''
            for page in reader.pages:
                text += page.extract_text()
            
            # Check if text is extracted
            if text.strip():
                # Move to pdf_TextMachina
                os.rename(filepath, os.path.join(pdf_textmachina_dir, filename))
            else:
                # Move to pdf_OCRobot
                os.rename(filepath, os.path.join(pdf_ocrobot_dir, filename))


# Run OCRobot/main.py
ocrobot_script = os.path.join(base_dir, 'OCRobot', 'main.py')
subprocess.run(['python', ocrobot_script])

# Run TextMachina/main.py
textmachina_script = os.path.join(base_dir, 'TextMachina', 'main.py')
subprocess.run(['python', textmachina_script])