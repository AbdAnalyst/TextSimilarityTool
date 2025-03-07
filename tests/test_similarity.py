"""
Unit Tests for Text Similarity Functions

Author: Your Name
Repository: https://github.com/your-username/TextSimilarityTool

Purpose:
This script tests the functions in `stored_procedure.py`
to ensure accuracy in text similarity calculations, data loading, and file generation.

Run Tests:
- If pytest is not installed, it will be automatically installed.
- Run tests using: pytest test_similarity.py
"""

import os
import subprocess
import sys

# ========================================================
# 1) Ensure pytest is Installed
# ========================================================
try:
    import pytest
except ImportError:
    print("pytest not found. Installing...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pytest"])
    import pytest

# Import functions to test
from stored_procedure import clean_and_tokenize, calculate_similarity, load_data, save_results_excel

# Sample texts for testing
text1 = "Heat stress due to high temperatures and humidity."
text2 = "High humidity can cause heat stress in workers."
text3 = "This is a completely unrelated sentence."

# ========================================================
# 2) Test Text Cleaning & Tokenization
# ========================================================
def test_clean_and_tokenize():
    tokens = clean_and_tokenize(text1)
    assert isinstance(tokens, set), "Output should be a set."
    assert "heat" in tokens, "'heat' should be a token."
    assert "stress" in tokens, "'stress' should be a token."
    assert "due" not in tokens, "Stopword 'due' should be removed."

# ========================================================
# 3) Test Similarity Calculation
# ========================================================
def test_calculate_similarity():
    sim1, common1 = calculate_similarity(text1, text2)
    sim2, common2 = calculate_similarity(text1, text3)

    assert 20.0 <= sim1 <= 50.0, "Expected similarity to be between 20-50%."
    assert common1, "Common words should exist between similar texts."
    assert sim2 == 0.0, "Unrelated texts should have 0% similarity."

# ========================================================
# 4) Test Data Loading
# ========================================================
@pytest.mark.parametrize("file_path, column_name", [
    ("riskAssessment_dummy_dataset.xlsx", "describeTheSpecificHazards"),
])
def test_load_data(file_path, column_name):
    data = load_data(file_path, column_name)
    assert isinstance(data, list), "Loaded data should be a list."
    assert len(data) > 0, "Dataset should not be empty."

# ========================================================
# 5) Test Excel File Generation
# ========================================================
def test_save_results_excel():
    test_data = [
        {"Text 1": text1, "Text 2": text2, "Similarity %": 35.0, "Common Words": "heat, stress, humidity"}
    ]
    output_file = "test_output.xlsx"
    
    save_results_excel(test_data, output_file)
    assert os.path.exists(output_file), "Excel file should be created."
    
    # Cleanup test file
    os.remove(output_file)

# ========================================================
# Run Tests
# ========================================================
if __name__ == "__main__":
    pytest.main()
