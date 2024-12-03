import pandas as pd
import os
import re
import shutil

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
    
    # Remplace les tirets par des espaces
    name = name.replace('-', ' ')
    
    # Supprime les caractères spéciaux tout en gardant les espaces
    name = re.sub(r'[^A-Z0-9\s]', '', name)
    
    # Normalise les espaces (supprime les espaces multiples)
    return ' '.join(word for word in name.split())

def get_name_variations(nom, nom_ancien, prenom):
    variations = set()
    noms = [normalize_name(nom)]
    
    # Gestion du nom ancien qui peut être un deuxième nom
    if nom_ancien and nom_ancien != nom:
        ancien_parts = normalize_name(nom_ancien).split()
        for part in ancien_parts:
            if part not in noms:
                noms.append(part)
    
    prenom_norm = normalize_name(prenom)
    
    # Génère toutes les combinaisons possibles
    for current_nom in noms:
        variations.add(f"{current_nom} {prenom_norm}")
        variations.add(f"{prenom_norm} {current_nom}")
        
        # Si le nom contient plusieurs parties
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
    target_directory = os.path.join(source_directory, 'fichiers_IPP')
    os.makedirs(target_directory, exist_ok=True)
    
    copied_count = 0
    errors = []
    
    for filename in os.listdir(source_directory):
        if not filename.endswith('.xlsx'):
            continue
            
        base_name = os.path.splitext(filename)[0]
        normalized_variations = [
            normalize_name(base_name),
            ' '.join(reversed(normalize_name(base_name).split()))
        ]
        
        matched = False
        for normalized_name in normalized_variations:
            if normalized_name in patient_dict:
                new_name = f"{patient_dict[normalized_name]}.xlsx"
                source_path = os.path.join(source_directory, filename)
                target_path = os.path.join(target_directory, new_name)
                
                try:
                    if os.path.exists(target_path):
                        errors.append(f"Erreur: {new_name} existe déjà pour {filename}")
                        continue
                        
                    shutil.copy2(source_path, target_path)
                    copied_count += 1
                    print(f"Copié: {filename} -> {new_name}")
                    matched = True
                    break
                except Exception as e:
                    errors.append(f"Erreur lors de la copie de {filename}: {str(e)}")
        
        if not matched:
            errors.append(f"Pas de correspondance trouvée pour: {filename}")
    
    return copied_count, errors

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
    copied_count, errors = process_files(directory, patient_dict)
    
    print(f"\nRésumé: {copied_count} fichiers copiés")
    if errors:
        print("\nErreurs:")
        for error in errors:
            print(error)

if __name__ == "__main__":
    main()