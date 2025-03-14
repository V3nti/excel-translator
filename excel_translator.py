import pandas as pd
import os
from deep_translator import GoogleTranslator
import time
import argparse
import sys

# Dictionary mapping user-friendly language names to language codes
LANGUAGE_MAP = {
    "english": "en",
    "french": "fr",
    "spain": "es",
    "spanish": "es",
    "german": "de",
    "italy": "it",
    "italian": "it",
    "dutch": "nl",
    "portuguese": "pt",
    "russian": "ru",
    "chinese": "zh-CN",
    "japanese": "ja",
    "korean": "ko",
    "arabic": "ar",
    "hindi": "hi",
    "turkish": "tr",
    "greek": "el",
    "polish": "pl",
    "vietnamese": "vi",
    "thai": "th",
    "swedish": "sv",
    "danish": "da",
    "finnish": "fi",
    "norwegian": "no",
    "czech": "cs",
    "romanian": "ro",
    "hungarian": "hu",
    "bulgarian": "bg",
    "ukrainian": "uk",
    "croatian": "hr",
    "slovak": "sk",
    "indonesia": "id",
    "indonesian": "id",
    "malay": "ms",
    "hebrew": "he",
    "latin": "la",
    # Country codes also map to language codes
    "fr": "fr",
    "en": "en",
    "eng": "en",
    "es": "es",
    "de": "de",
    "it": "it",
    "nl": "nl",
    "pt": "pt",
    "ru": "ru",
    "cn": "zh-CN",
    "jp": "ja",
    "kr": "ko",
    "ar": "ar",
    "in": "hi",
    "tr": "tr",
    "gr": "el",
    "pl": "pl",
    "vn": "vi",
    "th": "th",
    "se": "sv",
    "dk": "da",
    "fi": "fi",
    "no": "no",
    "cz": "cs",
    "ro": "ro",
    "hu": "hu",
    "bg": "bg",
    "ua": "uk",
    "hr": "hr",
    "sk": "sk",
    "id": "id",
    "my": "ms",
    "il": "he",
}

def normalize_language(lang):
    """
    Normalize language input to a valid language code.
    Returns the language code or None if not recognized.
    """
    lang = lang.lower().strip()
    return LANGUAGE_MAP.get(lang)

def translate_excel(input_file, target_languages, output_location=None):
    """
    Translates an Excel file into multiple languages and saves the translated files
    in the specified location with naming pattern: original_filename_language.xlsx
    
    Args:
        input_file (str): Path to the input Excel file
        target_languages (list): List of target language names or codes
        output_location (str): Path where translated files will be saved (defaults to same as input file)
    """
    try:
        # Read the Excel file
        print(f"Reading Excel file: {input_file}")
        try:
            df = pd.read_excel(input_file)
        except FileNotFoundError:
            print(f"Error: Input file '{input_file}' not found.")
            return
        except Exception as e:
            print(f"Error: Failed to read Excel file: {e}")
            return
        
        # Get the base filename without extension
        base_filename = os.path.splitext(os.path.basename(input_file))[0]
        
        # Set output location
        if output_location is None:
            # If not specified, use the same folder as the input file
            output_location = os.path.dirname(input_file)
            if not output_location:  # If input_file doesn't contain a directory path
                output_location = '.'
        
        # Create output folder if it doesn't exist
        try:
            if not os.path.exists(output_location):
                os.makedirs(output_location)
                print(f"Created output directory: {output_location}")
        except Exception as e:
            print(f"Error: Could not create output directory: {e}")
            return
        
        # Normalize and validate target languages
        valid_languages = []
        for lang in target_languages:
            lang_code = normalize_language(lang)
            if lang_code:
                valid_languages.append((lang, lang_code))
            else:
                print(f"Warning: Unrecognized language '{lang}' - skipping")
        
        if not valid_languages:
            print("Error: No valid languages specified for translation.")
            return
        
        # Process each target language
        for lang_name, lang_code in valid_languages:
            print(f"\nTranslating to {lang_name} ({lang_code})...")
            
            # Create a copy of the original dataframe
            translated_df = df.copy()
            
            try:
                # Initialize translator
                translator = GoogleTranslator(source='auto', target=lang_code)
                
                # Track progress
                total_cells = 0
                translated_cells = 0
                
                # Identify translatable columns first
                translatable_columns = []
                for col in translated_df.columns:
                    if translated_df[col].dtype == 'object':
                        translatable_columns.append(col)
                
                if not translatable_columns:
                    print("  No text columns found to translate.")
                    continue
                
                # Count total cells to translate for progress reporting
                for col in translatable_columns:
                    for idx in range(len(translated_df)):
                        cell_value = translated_df.at[idx, col]
                        if isinstance(cell_value, str) and cell_value.strip():
                            total_cells += 1
                
                # Translate each string column in the dataframe
                for col in translatable_columns:
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
                                translated_cells += 1
                                
                                # Report progress
                                if total_cells > 0:
                                    progress = (translated_cells / total_cells) * 100
                                    print(f"\r  Progress: {translated_cells}/{total_cells} cells ({progress:.1f}%)", end="")
                                
                                # Add a small delay to avoid hitting rate limits
                                time.sleep(0.2)
                            except Exception as e:
                                print(f"\n    Error translating '{cell_value}': {e}")
                
                print()  # New line after progress reporting
                
                # Save the translated dataframe to a new Excel file
                output_file = os.path.join(output_location, f"{base_filename}_{lang_code}.xlsx")
                try:
                    translated_df.to_excel(output_file, index=False)
                    print(f"Saved translated file: {output_file}")
                except Exception as e:
                    print(f"Error saving file {output_file}: {e}")
            
            except Exception as e:
                print(f"Error during translation to {lang_name}: {e}")
    
    except KeyboardInterrupt:
        print("\nTranslation interrupted by user.")
    except Exception as e:
        print(f"Unexpected error: {e}")

def main():
    # Set up command line arguments
    parser = argparse.ArgumentParser(description='Translate Excel files to multiple languages')
    parser.add_argument('input_file', help='Path to the input Excel file')
    parser.add_argument('--languages', '-l', nargs='+', required=True,
                        help='Target languages (e.g., english, french, es, de)')
    parser.add_argument('--output', '-o', 
                        help='Output location for translated files (default: same as input file)')
    parser.add_argument('--list-languages', action='store_true',
                        help='Display a list of supported language names and codes')
    
    args = parser.parse_args()
    
    # Show language list if requested
    if args.list_languages:
        print("Supported language names and codes:")
        languages = sorted(set(LANGUAGE_MAP.keys()))
        for lang in languages:
            code = LANGUAGE_MAP[lang]
            print(f"  {lang} -> {code}")
        return
    
    # Run the translation
    translate_excel(args.input_file, args.languages, args.output)

if __name__ == "__main__":
    main()
