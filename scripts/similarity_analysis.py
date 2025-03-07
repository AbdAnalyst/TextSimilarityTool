"""
Text Similarity Analysis Script

Author: Abdullah Choudhry
Repository: https://github.com/AbdAnalyst/TextSimilarityTool

Purpose:
This script analyzes text similarity by:
- Removing stopwords and punctuation.
- Tokenizing and comparing texts using a bag-of-words approach.
- Calculating similarity percentages.
- Highlighting common words in HTML.
- Saving only the first 20 results in both HTML and Excel.

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
import openpyxl
from nltk.corpus import stopwords
import nltk

# Ensure NLTK stopwords are available
nltk.download('stopwords')
stop_words = set(stopwords.words('english'))

# Load data from CSV
dataset_filename = "VariedRiskAssessmentDataset.csv"
column_to_analyze = "describeTheSpecificHazards"

df = pd.read_csv(dataset_filename)

if column_to_analyze not in df.columns:
    raise ValueError(f"Column '{column_to_analyze}' not found in the dataset.")

text_data = df[column_to_analyze].dropna().tolist()

# Function to clean and tokenize text
def clean_and_tokenize(text):
    text = text.lower()
    tokens = re.findall(r"[a-z0-9]+", text)
    return set(t for t in tokens if t not in stop_words)

# Function to calculate text similarity
def bag_of_words_similarity(text1, text2):
    words1, words2 = clean_and_tokenize(text1), clean_and_tokenize(text2)
    common_words = words1.intersection(words2)
    avg_size = (len(words1) + len(words2)) / 2.0
    similarity_percentage = (len(common_words) / avg_size) * 100 if avg_size else 0.0
    return similarity_percentage, common_words

# Function to highlight common words in HTML
def highlight_common_words(text, common_words, color="#FFD54F"):
    tokens = re.findall(r"\S+", text)
    return " ".join(
        f"<span style='background-color:{color}'>{token}</span>" if re.sub(r'\W+', '', token.lower()) in common_words else token
        for token in tokens
    )

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
            "Common Words": ", ".join(sorted(common_words)) if common_words else "None",
        })

# Keep only the first 20 rows for HTML and Excel output (without sorting)
top_comparisons = comparisons[:20]

# Generate HTML report
html_rows = []
for comp in top_comparisons:
    highlighted_text1 = highlight_common_words(comp["Text 1"], set(comp["Common Words"].split(", ")), "#FFD54F")
    highlighted_text2 = highlight_common_words(comp["Text 2"], set(comp["Common Words"].split(", ")), "#81C784")

    html_rows.append(f"""
        <tr>
            <td>{comp['Text 1']}</td>
            <td>{comp['Text 2']}</td>
            <td style="text-align:center;">{comp['Similarity %']:.2f}%</td>
            <td>{comp['Common Words']}</td>
            <td>{highlighted_text1}</td>
            <td>{highlighted_text2}</td>
        </tr>
    """)

html_output = f"""
<h2>First 20 Text Similarity Analysis Results</h2>
<p>Displaying the first 20 text comparisons as they appear.</p>
<table border="1" style="border-collapse:collapse;width:100%;">
  <thead>
    <tr>
      <th>Original Text 1</th>
      <th>Original Text 2</th>
      <th>Similarity %</th>
      <th>Common Words</th>
      <th>Highlighted Text 1</th>
      <th>Highlighted Text 2</th>
    </tr>
  </thead>
  <tbody>
    {''.join(html_rows)}
  </tbody>
</table>
"""

html_filename = "first_20_overlap_results.html"
with open(html_filename, "w", encoding="utf-8") as f:
    f.write(html_output)

print(f"HTML report (First 20) saved to {html_filename}")

# Save first 20 results in Excel
excel_filename = "first_20_overlap_results.xlsx"
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "First 20 Text Similarity Results"
ws.append(["Text 1", "Text 2", "Similarity %", "Common Words"])

for comp in top_comparisons:
    ws.append([comp["Text 1"], comp["Text 2"], comp["Similarity %"], comp["Common Words"]])

wb.save(excel_filename)
print(f"First 20 results saved to {excel_filename}")
