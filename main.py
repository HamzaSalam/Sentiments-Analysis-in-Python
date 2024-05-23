import os
import pandas as pd
from openpyxl import Workbook
from roman_urdu_sentiments import positive_words, negative_words

def get_sentiment(text):
    """Determine if the sentiment of the text is positive or negative."""
    words = text.lower().split()
    positive_count = sum(1 for word in words if word in positive_words)
    negative_count = sum(1 for word in words if word in negative_words)
    
    if positive_count > negative_count:
        return 'Positive'
    elif negative_count > positive_count:
        return 'Negative'
    else:
        return 'Neutral'

def save_to_excel(data, filename='sentiments.xlsx'):
    """Save the list of dictionaries to an Excel file."""
    if not os.path.exists(filename):
        # Create a new Excel file if it doesn't exist
        wb = Workbook()
        wb.save(filename)
    
    existing_data = pd.read_excel(filename) if os.path.exists(filename) else pd.DataFrame()
    
    df = pd.DataFrame(data)
    combined_df = pd.concat([existing_data, df], ignore_index=True)
    
    with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
        combined_df.to_excel(writer, index=False)

def main():
    user_inputs = []
    print("Enter 'exit' to quit and save")
    while True:
        user_input = input("Enter something: ")
        if user_input.lower() == 'exit':
            break
        sentiment = get_sentiment(user_input)
        print(f"Sentiment: {sentiment}")
        user_inputs.append({'Text': user_input, 'Sentiment': sentiment})

    if user_inputs:
        save_to_excel(user_inputs)
        print("Data saved to sentiments.xlsx")

if __name__ == "__main__":
    main()
