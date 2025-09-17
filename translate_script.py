import pandas as pd
from deep_translator import GoogleTranslator, DeeplTranslator
import os

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
        
        #translator = DeeplTranslator(source=source_lang, target=target_lang)

        # Create a new column for the translated text
        translated_column_name = f"{column_to_translate}_{target_lang}"
        df[translated_column_name] = ""  # Initialize with empty strings
        
        print(f"Translating column '{column_to_translate}'...")
        
        # Iterate through the column and translate each cell
        for row_num, (index, row) in enumerate(df.iterrows(), start=1):
            text_to_translate = row[column_to_translate]
            
            # Check if the cell contains a string
            if isinstance(text_to_translate, str) and text_to_translate.strip() != "":
                translated_text = translator.translate(text_to_translate)
                df.at[index, translated_column_name] = translated_text
            else:
                # If the cell is empty or not a string, keep it as is
                df.at[index, translated_column_name] = text_to_translate
                
            print(f"Row {row_num}: Translated '{text_to_translate}' to '{translated_text}'")

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
    
    input_file = "excel_files/nl_translations.xlsx"
    sheet = "Sheet1"
    text_column = "en"  # Name of the column with text to translate
    
    # You can find language codes (ISO 639-1) for many languages online.
    source_language = "en"  # English
    target_language = "nl"  # Dutch

    # --- Run the translation script ---
    translate_excel_file(input_file, sheet, text_column, source_language, target_language)