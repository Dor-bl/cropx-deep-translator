import pandas as pd
import os
import time
from dotenv import load_dotenv
from openai_translator import initialize_openai, translate_with_openai

# Load environment variables from .env file
load_dotenv()

# Initialize the OpenAI API key
initialize_openai()

def translate_excel_file(file_path: str, sheet_name: str, column_to_translate: str, source_lang: str, target_lang: str) -> None:
    """
    Translates a specific column in an Excel file and saves the result to a new file.
    """
    try:
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Create a new column for the translated text
        translated_column_name = f"{column_to_translate}_{target_lang}"
        
        print(f"Translating column '{column_to_translate}' using OpenAI...")

        # Start timer
        start_time = time.time()

        # Apply the custom translation function to the column
        df[translated_column_name] = df[column_to_translate].apply(
            lambda x: translate_with_openai(x, source_lang, target_lang)
        )

        # End timer and calculate duration
        end_time = time.time()
        duration = end_time - start_time
        print(f"Translation process took {duration:.2f} seconds.")

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
    input_file = "excel_files/nl_translations.xlsx"
    sheet = "Sheet1"
    text_column = "en"
    source_language = "English"
    target_language = "Dutch"

    # --- Run the translation script ---
    translate_excel_file(input_file, sheet, text_column, source_language, target_language)