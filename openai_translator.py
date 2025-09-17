import openai
import os

def initialize_openai():
    """Initializes the OpenAI API key from environment variables and checks it."""
    openai.api_key = os.getenv("OPENAI_API_KEY")
    if not openai.api_key or openai.api_key == "your_api_key_here":
        raise ValueError("Please set your OPENAI_API_KEY in the .env file.")

def translate_with_openai(text: str, source_lang: str, target_lang: str) -> str:
    """
    Translates text using a custom implementation of the OpenAI API.
    """
    if not text or not isinstance(text, str):
        return text

    try:
        # Construct a clear prompt for the model
        prompt = f"Translate the following text from {source_lang} to {target_lang}. Only return the translated text, without any additional comments or explanations: '{text}'"

        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",  # You can swap this with "gpt-4" or other models
            messages=[
                {"role": "system", "content": "You are a professional translator."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,  # Lower temperature for more deterministic translations
            max_tokens=len(text) * 2  # A reasonable limit for the output
        )
        content = response.choices[0].message.content
        if content is not None:
            translated_text = content.strip()
            # Clean up potential quote marks from the model's response
            if translated_text.startswith(('"', "'")) and translated_text.endswith(('"', "'")):
                translated_text = translated_text[1:-1]
            return translated_text
        else:
            print("Warning: OpenAI response content is None.")
            return text
    except Exception as e:
        print(f"An error occurred during OpenAI translation: {e}")
        return text # Return original text on failure