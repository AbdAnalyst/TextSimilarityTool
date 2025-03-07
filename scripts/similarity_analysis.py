"""
Text Similarity Analysis Script

Author: Your Name
Repository: https://github.com/your-username/TextSimilarityTool

Purpose:
This script analyzes text similarity by:
- Removing stopwords and punctuation.
- Tokenizing and comparing texts using a bag-of-words approach.
- Calculating similarity percentages.
- Highlighting common words in HTML.
- Exporting results to an Excel file.

Usage:
1. Install dependencies:
   pip install pandas openpyxl nltk ipython
2. Download NLTK stopwords:
   import nltk
   nltk.download('stopwords')
3. Place the dataset (CSV file) in the same folder.
4. Run the script:
   python similarity_analysis.py
"""

import re
import pandas as pd
from IPython.display import display, HTML
import openpyxl
from openpyxl.styles import Font
from nltk.corpus import stopwords
import nltk

# Ensure NLTK stopwords are available
nltk.download('stopwords')
stop_words = set(stopwords.words('english'))

# Load data from CSV
dataset_filename = "VariedRiskAssessmentDataset.csv"  # Updated to CSV file
column_to_analyze = "describeTheSpecificHazards"  # Ensure this column exists in CSV

# Read the dataset
df = pd.read_csv(dataset_filename)

# Ensure the column exists
if column_to_analyze not in df.columns:
    raise ValueError(f"Column '{column_to_analyze}' not found in the dataset.")

# Extract non-empty text rows
text_data = df[column_to_analyze].dropna().tolist()

# Function to clean and tokenize text
def clean_and_tokenize(text):
    """
    Cleans text by removing punctuation, converting to lowercase,
    and filtering out stopwords.
    """
    text = text.lower()
    tokens = re.findall(r"[a-z0-9]+", text)
    return set(t for t in tokens if t not in stop_words)

# Function to calculate text similarity
def bag_of_words_similarity(text1, text2):
    """
    Calculates similarity percentage based on shared words.
    Ignores punctuation and stopwords.
    """
    words1, words2 = clean_and_tokenize(text1), clean_and_tokenize(text2)
    common_words = words1.intersection(words2)
    avg_size = (len(words1) + len(words2)) / 2.0
    similarity_percentage = (len(common_words) / avg_size) * 100 if avg_size else 0.0
    return similarity_percentage, common_words

# Run comparisons
comparisons = []
for i in range(len(text_data)):
    for j in range(i + 1, len(text_data)):
        text1, text2 = text_data[i], text_data[j]
        similarity, common_words = bag_of_words_similarity(text1, text2)
        comparisons.append({
            "Text 1": text1,
            "Text 2": text2,
            "Similarity %": round(similarity, 2),
            "Common Words": ", ".join(sorted(common_words)) if common_words else "None"
        })

# Save results as Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Text Similarity Results"
ws.append(["Text 1", "Text 2", "Similarity %", "Common Words"])

for comp in comparisons:
    ws.append([comp["Text 1"], comp["Text 2"], comp["Similarity %"], comp["Common Words"]])

excel_filename = "overlap_results.xlsx"
wb.save(excel_filename)
print(f"Excel file saved to {excel_filename}")
