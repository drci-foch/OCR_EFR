import pandas as pd
import os
import shutil

def clean_ipp(ipp):
    # Convert to string and remove all spaces
    return str(ipp).strip().replace(' ', '')

def process_files(source_directory, mapping_df):
    # Create output directory
    target_directory = os.path.join(source_directory, 'renamed_files')
    os.makedirs(target_directory, exist_ok=True)
    
    results = []
    copied_count = 0
    
    # Convert mapping DataFrame to dictionary for faster lookup
    # Clean IPP values while creating the dictionary
    mapping_dict = dict(zip(mapping_df['Excel_Filename'], 
                          [clean_ipp(ipp) for ipp in mapping_df['IPP']]))
    
    # Process each Excel file in the directory
    for filename in os.listdir(source_directory):
        if not filename.endswith('.xlsx'):
            continue
        
        # Look up IPP from mapping
        ipp = mapping_dict.get(filename, 'Non trouvé')
        
        results.append({
            'Excel_Filename': filename,
            'IPP': ipp
        })
        
        # Copy and rename file if IPP was found
        if ipp != 'Non trouvé':
            source_path = os.path.join(source_directory, filename)
            new_filename = f"{ipp}.xlsx"
            target_path = os.path.join(target_directory, new_filename)
            
            try:
                shutil.copy2(source_path, target_path)
                copied_count += 1
                print(f"Copié: {filename} -> {new_filename}")
            except Exception as e:
                print(f"Erreur lors de la copie de {filename}: {str(e)}")
    
    return results, copied_count

def main():
    # Read the mapping file
    try:
        mapping_df = pd.read_excel('exact_mapping.xlsx')  # Replace with your mapping file name
        print("Fichier de mapping lu avec succès")
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier de mapping: {str(e)}")
        return
    
    directory = '.'
    results, copied_count = process_files(directory, mapping_df)
    
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
    print(f"Fichiers copiés et renommés: {copied_count}")
    print(f"\nFichiers renommés sauvegardés dans: {os.path.join(directory, 'renamed_files')}")
    print(f"Rapport de correspondance sauvegardé dans: {output_filename}")

if __name__ == "__main__":
    main()