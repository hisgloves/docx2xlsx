# DOCX to Excel Translation Table Script

This repository contains a Python script to process DOCX files in a folder, extract text, split it into sentences or lines, and create corresponding Excel files with translation tables. The script ensures that each sentence or line is placed in a separate cell in the Excel file, facilitating easier translation.

## Features

- Processes all DOCX files in the current folder.
- Extracts text from DOCX files and splits it into sentences and lines.
- Creates Excel files with two columns: "Original" and "Translation".
- Skips processing if the corresponding Excel file already exists.
- Counts characters in both DOCX and Excel files to ensure accuracy, excluding non-content characters (disabled by default).
- Notifies if there is a mismatch in character counts (disabled by default).

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/docx-to-excel-translation.git
   cd docx-to-excel-translation
2. **Install the required libraries:**
Make sure you have pip installed. Then run:
   ```bash
   pip install python-docx nltk openpyxl
3. **Download NLTK resources:**
The script requires NLTK resources for sentence tokenization. You can download them using the following command:
   ```bash
   import nltk
   nltk.download('punkt')
## Usage
1. Place your DOCX files:
Place all the DOCX files you want to process in the same directory as the script.

2. Run the script:
Execute the script using Python:
   ```bash
   python docx2xlsx.py
3. Output:
The script will create corresponding Excel files for each DOCX file in the same directory. Each Excel file will have the name format **original_name_translate.xlsx**.
## License
This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- [python-docx](https://python-docx.readthedocs.io/)
- [NLTK](https://www.nltk.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- This script and README were generated with the assistance of [OpenAI's ChatGPT 4o](https://openai.com/chatgpt).
