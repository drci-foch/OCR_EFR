import pandas as pd
import os
import re

def normalize_name(name):
    name = name.upper().strip()
    accents = {
        'É': 'E', 'È': 'E', 'Ê': 'E', 'Ë': 'E',
        'À': 'A', 'Â': 'A', 'Ä': 'A',
        'Î': 'I', 'Ï': 'I',
        'Ô': 'O', 'Ö': 'O',
        'Ù': 'U', 'Û': 'U', 'Ü': 'U',
        'Ç': 'C'
    }
    for acc, repl in accents.items():
        name = name.replace(acc, repl)
    
    name = name.replace('-', ' ')
    name = re.sub(r'[^A-Z0-9\s]', '', name)
    return ' '.join(word for word in name.split())

def get_name_variations(nom, nom_ancien, prenom):
    variations = set()
    noms = [normalize_name(nom)]
    
    if nom_ancien and nom_ancien != nom:
        ancien_parts = normalize_name(nom_ancien).split()
        for part in ancien_parts:
            if part not in noms:
                noms.append(part)
    
    prenom_norm = normalize_name(prenom)
    
    for current_nom in noms:
        variations.add(f"{current_nom} {prenom_norm}")
        variations.add(f"{prenom_norm} {current_nom}")
        
        nom_parts = current_nom.split()
        if len(nom_parts) > 1:
            for part in nom_parts:
                variations.add(f"{part} {prenom_norm}")
                variations.add(f"{prenom_norm} {part}")
    
    return {v for v in variations if v.strip()}

def create_patient_dict(df):
    patient_dict = {}
    
    for _, row in df.iterrows():
        nip = row['NIP']
        variations = get_name_variations(row['nom'], row['nom.ancienne'], row['prenom'])
        
        for variation in variations:
            patient_dict[variation] = nip
    
    return patient_dict

def process_files(source_directory, patient_dict):
    results = []
    
    for filename in os.listdir(source_directory):
        if not filename.endswith('.xlsx'):
            continue
            
        base_name = os.path.splitext(filename)[0]
        normalized_variations = [
            normalize_name(base_name),
            ' '.join(reversed(normalize_name(base_name).split()))
        ]
        
        matched_ipp = None
        for normalized_name in normalized_variations:
            if normalized_name in patient_dict:
                matched_ipp = patient_dict[normalized_name]
                break
        
        results.append({
            'Excel_Filename': filename,
            'IPP': matched_ipp if matched_ipp else 'Non trouvé'
        })
    
    return results

def main():
    encodings = ['latin1', 'utf-8', 'iso-8859-1', 'cp1252']
    df = None
    
    for encoding in encodings:
        try:
            df = pd.read_csv('list IPP NATT 18 11 2024.csv', sep=';', encoding=encoding)
            print(f"Fichier lu avec l'encodage: {encoding}")
            break
        except UnicodeDecodeError:
            continue
    
    if df is None:
        print("Impossible de lire le fichier CSV")
        return
        
    patient_dict = create_patient_dict(df)
    directory = '.'
    results = process_files(directory, patient_dict)
    
    # Create DataFrame and save to Excel
    results_df = pd.DataFrame(results)
    output_filename = 'matching_results.xlsx'
    results_df.to_excel(output_filename, index=False)
    
    # Print summary
    total_files = len(results)
    matched_files = sum(1 for r in results if r['IPP'] != 'Non trouvé')
    print(f"\nRésumé:")
    print(f"Total fichiers Excel: {total_files}")
    print(f"Fichiers avec IPP trouvé: {matched_files}")
    print(f"Fichiers sans correspondance: {total_files - matched_files}")
    print(f"\nRésultats sauvegardés dans: {output_filename}")

if __name__ == "__main__":
    main()