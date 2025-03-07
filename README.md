# Text Similarity Tool

🚀 Welcome to **TextSimilarityTool** – a powerful Python-based tool for analyzing text similarity!

📌 About This Project

TextSimilarityTool is designed to compare textual data by removing stopwords and punctuation, calculating word overlap, and highlighting similarities in an easy-to-read format. It’s useful for risk assessments, duplicate detection, content comparison, and other NLP-related tasks.

----

## ⚙️ Project Requirements

### 🎯 Objective

The goal of this project is to provide a simple yet effective text similarity detection tool that can:
	• Analyze textual descriptions to find common words.
	• Ignore stopwords and punctuation for more accurate similarity results.
	• Provide visual and tabular output in both HTML and Excel formats.
	• Support large datasets efficiently.

#### 📌 Specification

This project follows these specifications:
	• Uses Python 3.7+ for text analysis.
	• Applies NLP to process and clean text.
	• Bag-of-Words approach is used to compute text similarity.
	• Generates reports in:
	• HTML: Highlights common words for easy readability.
	• Excel: Stores results with partial font coloring for common words.
	• Works on any dataset with textual descriptions.

#### 🔧 Dependencies

To run this project, you need the following Python libraries:
	•	pandas – For data handling
	•	openpyxl – For Excel file creation
	•	nltk – For Natural Language Processing (stopword removal)
	•	re – Regular expressions for text cleaning
	•	IPython – For displaying formatted output in Jupyter Notebook

You can install all dependencies with:

pip install pandas openpyxl nltk ipython

Additionally, download the stopwords dataset for NLTK:

import nltk
nltk.download('stopwords')

------- 

#### 🚀 Getting Started

**Installation**
1.      Clone this repository:

git clone https://github.com/your-username/TextSimilarityTool.git
cd TextSimilarityTool

2.	Install required dependencies (as mentioned above).

**Usage**
• Place your dataset in the project folder (or use the provided dummy dataset).
• Run the script to analyze text similarity and generate reports in HTML and Excel.
• Open the HTML file for a visual representation of word overlap.

python similarity_analysis.py

⸻

## 📝 Contributing

Contributions are welcome! Feel free to submit issues or pull requests to improve the tool.

## 📄 License

This project is open-source and available under the MIT License.
