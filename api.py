import openai
import pandas as pd

class OpenAI_API:
    def __init__(self, api_key):
        self.api_key = api_key

    def __enter__(self):
        openai.api_key = self.api_key

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

def open_ai_response(question: str, dataframe: pd.DataFrame, sheet_name: str) -> str:
    chunk_size = 50000
    chunks = [dataframe[i:i+chunk_size] for i in range(0, len(dataframe), chunk_size)]
    result = ""

    for chunk in chunks:
        messages = [
            {"role": "system", "content": "Assistant is a large language model trained by OpenAI."},
            {"role": "user", "content": f"Sheet: {sheet_name}\n{question}\n{chunk.to_string()}"}
        ]

        response = openai.ChatCompletion.create(
            model="gpt-4-turbo",
            messages=messages,
            max_tokens=1000,
        )

        answer = response.choices[0].message.content.strip()
        result += answer

    return result


