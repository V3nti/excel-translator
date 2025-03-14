import requests
import json
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

def translate_text(text, target_language):
    """
    Translates the given text to the specified target language.
    
    Args:
        text (str): The text to translate
        target_language (str): The language code to translate to (e.g., 'es' for Spanish, 'fr' for French)
    
    Returns:
        str: The translated text
    """
    # You would need to set up your own API key for a translation service
    # This example uses a hypothetical API_KEY that should be stored in your .env file
    api_key = os.getenv('TRANSLATION_API_KEY')
    
    if not api_key:
        return "Error: API key not found. Please set the TRANSLATION_API_KEY in your .env file."
    
    # Using a generic translation API endpoint
    url = "https://api.translation-service.com/v2/translate"
    
    payload = {
        "q": text,
        "target": target_language,
        "format": "text"
    }
    
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    
    try:
        response = requests.post(url, headers=headers, data=json.dumps(payload))
        response.raise_for_status()  # Raise an exception for HTTP errors
        
        result = response.json()
        return result.get("translatedText", "Translation failed")
    
    except requests.exceptions.RequestException as e:
        return f"Error: {str(e)}"

def main():
    """
    Main function to get user input and display the translation.
    """
    print("Text Translation Tool")
    print("---------------------")
    
    text = input("Enter the text to translate: ")
    target_language = input("Enter the target language code (e.g., 'es' for Spanish, 'fr' for French): ")
    
    print("\nTranslating...")
    translated_text = translate_text(text, target_language)
    
    print(f"\nOriginal: {text}")
    print(f"Translated ({target_language}): {translated_text}")

if __name__ == "__main__":
    main()
