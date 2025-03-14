import pandas as pd
import os
from deep_translator import GoogleTranslator
import time
import argparse

def translate_excel(input_file, target_languages):
    """
    Translates an Excel file into multiple languages and saves the translated files
    in the same folder as the original with naming pattern: original_filename_language.xlsx
    
    Args:
        input_file (str): Path to the input Excel file
        target_languages (list): List of target language codes (e.g., 'fr', 'es', 'de')
    """
    # Read the Excel file
    print(f"Reading Excel file: {input_file}")
    df = pd.read_excel(input_file)
    
    # Get the folder path and filename without extension
    folder_path = os.path.dirname(input_file)
    if not folder_path:  # If input_file doesn't contain a directory path
        folder_path = '.'
    
    base_filename = os.path.splitext(os.path.basename(input_file))[0]
    
    # Process each target language
    for lang in target_languages:
        print(f"\nTranslating to {lang}...")
        
        # Create a copy of the original dataframe
        translated_df = df.copy()
        
        # Initialize translator
        translator = GoogleTranslator(source='auto', target=lang)
        
        # Translate each string column in the dataframe
        for col in translated_df.columns:
            # Only translate string columns
            if translated_df[col].dtype == 'object':
                print(f"  Translating column: {col}")
                
                # Translate each cell in the column
                for idx in range(len(translated_df)):
                    cell_value = translated_df.at[idx, col]
                    
                    # Only translate string values that are not empty
                    if isinstance(cell_value, str) and cell_value.strip():
                        try:
                            # Translate the text
                            translated_text = translator.translate(cell_value)
                            translated_df.at[idx, col] = translated_text
                            
                            # Add a small delay to avoid hitting rate limits
                            time.sleep(0.2)
                        except Exception as e:
                            print(f"    Error translating '{cell_value}': {e}")
        
        # Save the translated dataframe to a new Excel file in the same folder
        output_file = os.path.join(folder_path, f"{base_filename}_{lang}.xlsx")
        translated_df.to_excel(output_file, index=False)
        print(f"Saved translated file: {output_file}")

if __name__ == "__main__":
    # Set up command line arguments
    parser = argparse.ArgumentParser(description='Translate Excel files to multiple languages')
    parser.add_argument('input_file', help='Path to the input Excel file')
    parser.add_argument('--languages', '-l', nargs='+', default=['fr', 'es', 'de'], 
                        help='Target language codes (default: fr, es, de)')
    
    args = parser.parse_args()
    
    # Run the translation
    translate_excel(args.input_file, args.languages)
