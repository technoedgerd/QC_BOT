import re
import pandas as pd

# Load the US to UK dictionary from CSV
def load_us_to_uk_dict(csv_path="us_to_uk_dictionary.csv"):
    df = pd.read_csv(csv_path)
    dictionary = dict(zip(df["US"], df["UK"]))
    return dictionary

# Replace US words with UK equivalents
def convert_us_to_uk(text, dictionary):
    for us, uk in dictionary.items():
        pattern = r'\b' + re.escape(us) + r'\b'
        text = re.sub(pattern, uk, text, flags=re.IGNORECASE)
    return text

# Detect presence of US spellings in the text
def detect_us_words(text, dictionary):
    found = []
    for us in dictionary.keys():
        pattern = r'\b' + re.escape(us) + r'\b'
        if re.search(pattern, text, flags=re.IGNORECASE):
            found.append(us)
    return found
