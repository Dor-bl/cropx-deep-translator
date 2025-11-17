import pandas as pd
from deep_translator import GoogleTranslator, DeeplTranslator, MicrosoftTranslator
import os
import time
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

def translate_excel_file(file_path: str, sheet_name: str, column_to_translate: str, source_lang: str, target_lang: str) -> None:
    """
    Translates a specific column in an Excel file and saves the result to a new file.

    Args:
        file_path (str): The path to the input Excel file.
        sheet_name (str): The name of the sheet to work with.
        column_to_translate (str): The name of the column containing text to translate.
        source_lang (str): The source language code (e.g., 'en' for English).
        target_lang (str): The target language code (e.g., 'es' for Spanish).
    """
    try:
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Initialize the translator
        translator = GoogleTranslator(source=source_lang, target=target_lang)

        # Get API key from environment variables
        chatgpt_api_key = os.getenv("CHATGPT_API_KEY")
        if not chatgpt_api_key or chatgpt_api_key == "your_api_key_here":
            raise ValueError("Please set your CHATGPT_API_KEY in the .env file.")

        #translator = MicrosoftTranslator(api_key=chatgpt_api_key, source=source_lang, target=target_lang)


        # Create a new column for the translated text
        translated_column_name = f"{target_lang}"
        
        print(f"Translating column '{column_to_translate}'...")
        print(f"Total rows to translate: {len(df)}")
        print("-" * 50)

        # Start timer
        start_time = time.time()

        # Create a list to store translations
        translations = []
        
        # Translate each cell with progress feedback
        for idx, cell_value in enumerate(df[column_to_translate], 1):
            if isinstance(cell_value, str) and cell_value.strip() != "":
                translated_text = translator.translate(cell_value)
                translations.append(translated_text)
                print(f"[{idx}/{len(df)}] Translated: {cell_value[:50]}{'...' if len(cell_value) > 50 else ''}")
            else:
                translations.append(cell_value)
                print(f"[{idx}/{len(df)}] Skipped empty cell")
        
        # Assign the translations to the new column
        df[translated_column_name] = translations

        # End timer and calculate duration
        end_time = time.time()
        duration = end_time - start_time
        print("-" * 50)
        print(f"Translation process took {duration:.2f} seconds.")
        print(f"Average time per row: {duration/len(df):.2f} seconds")

        # Define the output file name
        output_file_path = f"{os.path.splitext(file_path)[0]}_translated.xlsx"
        
        # Save the new DataFrame to an Excel file
        df.to_excel(output_file_path, index=False)
        
        print(f"\nTranslation complete! New file saved to '{output_file_path}'")
        
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
    except KeyError:
        print(f"Error: The column '{column_to_translate}' or sheet '{sheet_name}' was not found in the file.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    # --- Configuration ---
    # Make sure to replace these values with your specific file and language details.
    
    input_file = "excel_files/nl_mobile_translation.xlsx"
    sheet = "Sheet1"
    text_column = "en"  # Name of the column with text to translate
    
    # You can find language codes (ISO 639-1) for many languages online.
    source_language = "en"  # English
    target_language = "nl"  # Dutch

    # --- Run the translation script ---
    translate_excel_file(input_file, sheet, text_column, source_language, target_language)