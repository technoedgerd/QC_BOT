import re
import pandas as pd
from us2uk_QC import load_us_to_uk_dict, detect_us_words

# Load US to UK dictionary
us_uk_dict = load_us_to_uk_dict()

# Pattern to detect contractions with smart (’) and straight (') apostrophes
CONTRACTION_PATTERN = re.compile(
    r"\b(?:I['’]m|don['’]t|can['’]t|won['’]t|shouldn['’]t|could['’]ve|you['’]d|it['’]s|he['’]d|she['’]d|they['’]re|we['’]re|didn['’]t|doesn['’]t|isn['’]t|wasn['’]t|aren['’]t|weren['’]t|['’]re|['’]d|['’]ll|['’]ve|['’]m|['’]s|n['’]t)\b",
    re.IGNORECASE
)

def has_contraction(text):
    return bool(CONTRACTION_PATTERN.search(text))

def has_us_spelling(text):
    return detect_us_words(text, us_uk_dict)

def has_extra_spaces(text):
    return "  " in text

def has_ending_period(text):
    return text.strip().endswith(".")

def has_mid_sentence_period(text):
    return bool(re.search(r"(?<!^)\.(?!$)", text))

def run_text_rules_validation(df_qc):
    if "Extracted Text" not in df_qc.columns:
        return pd.DataFrame([{"Error": "Missing 'Extracted Text' column in QC sheet"}])

    result = []
    for _, row in df_qc.iterrows():
        text = str(row.get("Extracted Text", "")).strip()

        # Skip completely empty lines
        if not text:
            continue

        contraction_found = has_contraction(text)
        us_words = has_us_spelling(text)

        result.append({
            "File Name": row.get("File Name", ""),
            "Slide Number": row.get("Slide Number", ""),
            "Shape Name / Table Cell": row.get("Shape Name / Table Cell", ""),
            "Extracted Text": text,
            "Contraction Used": "Yes" if contraction_found else "",
            "US English Used": ", ".join(us_words) if us_words else "",
            "Extra Space": "Yes" if has_extra_spaces(text) else "",
            "Ending Period Used": "Yes" if has_ending_period(text) else "",
            "Mid Sentence Period Used": "Yes" if has_mid_sentence_period(text) else ""
        })

    return pd.DataFrame(result)

# Debug mode
if __name__ == "__main__":
    df = pd.read_excel("example_input_qc.xlsx")
    df_out = run_text_rules_validation(df)
    df_out.to_excel("text_rules_output.xlsx", index=False)
