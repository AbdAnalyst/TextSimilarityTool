"""
Stored Procedure for Text Similarity Analysis

Author: Your Name
Repository: https://github.com/your-username/TextSimilarityTool

Purpose:
This script provides modular functions for text similarity analysis, making it
easy to reuse in other scripts or integrate into larger applications.

Usage:
1. Import the required functions in another script:
   from stored_procedure import load_data, analyze_texts, save_results_excel
2. Run analysis using:
   text_data = load_data("riskAssessment_dummy_dataset.xlsx", "describeTheSpecificHazards")
   comparisons = analyze_texts(text_data)
   save_results_excel(comparisons, "output.xlsx")
"""

import re
import pandas as pd
import openpyxl
import nltk
from nltk.corpus import stopwords

# Ensure stopwords are available
nltk.download('stopwords')
stop_words = set(stopwords.words('english'))

# ========================================================
# 1) Load Data from Excel
# ========================================================
def load_data(file_path, column_name):
    """
    Loads text data from an Excel file.
    :param file_path: str - Path to the Excel file.
    :param column_name: str - Column containing text descriptions.
    :return: List of text rows.
    """
    df = pd.read_excel(file_path, engine="openpyxl")
    return df[column_name].dropna().tolist()

# ========================================================
# 2) Clean & Tokenize Text
# ========================================================
def clean_and_tokenize(text):
    """
    Cleans text by removing punctuation, converting to lowercase,
    and filtering out stopwords.
    :param text: str - Input text.
    :return: Set of cleaned tokens.
    """
    text = text.lower()
    tokens = re.findall(r"[a-z0-9]+", text)
    return set(t for t in tokens if t not in stop_words)

# ========================================================
# 3) Calculate Text Similarity
# ========================================================
def calculate_similarity(text1, text2):
    """
    Calculates similarity percentage based on shared words.
    :param text1: str - First text.
    :param text2: str - Second text.
    :return: (float, set) - Similarity percentage and common words.
    """
    words1, words2 = clean_and_tokenize(text1), clean_and_tokenize(text2)
    common_words = words1.intersection(words2)
    avg_size = (len(words1) + len(words2)) / 2.0
    similarity_percentage = (len(common_words) / avg_size) * 100 if avg_size else 0.0
    return similarity_percentage, common_words

# ========================================================
# 4) Compare Texts & Generate Results
# ========================================================
def analyze_texts(text_list):
    """
    Compares texts, calculates similarity, and formats output.
    :param text_list: List of text descriptions.
    :return: List of comparison results.
    """
    comparisons = []
    for i in range(len(text_list)):
        for j in range(i + 1, len(text_list)):
            text1, text2 = text_list[i], text_list[j]
            similarity, common_words = calculate_similarity(text1, text2)

            comparisons.append({
                "Text 1": text1,
                "Text 2": text2,
                "Similarity %": round(similarity, 2),
                "Common Words": ", ".join(sorted(common_words)) if common_words else "None",
            })
    return comparisons

# ========================================================
# 5) Save Results as Excel
# ========================================================
def save_results_excel(comparisons, filename="overlap_results.xlsx"):
    """
    Saves similarity results in an Excel file.
    :param comparisons: List of comparison results.
    :param filename: str - Output Excel filename.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Text Similarity Results"

    headers = ["Text 1", "Text 2", "Similarity %", "Common Words"]
    ws.append(headers)

    for comp in comparisons:
        ws.append([comp["Text 1"], comp["Text 2"], comp["Similarity %"], comp["Common Words"]])

    wb.save(filename)
    print(f"Excel file saved to {filename}")

# ========================================================
# 6) Run the Script (Only When Executed Directly)
# ========================================================
if __name__ == "__main__":
    dataset_filename = "riskAssessment_dummy_dataset.xlsx"
    column_to_analyze = "describeTheSpecificHazards"

    print("Loading dataset...")
    text_data = load_data(dataset_filename, column_to_analyze)

    print("Analyzing text similarity...")
    comparisons = analyze_texts(text_data)

    print("Saving results...")
    save_results_excel(comparisons)

    print("Process completed successfully!")
