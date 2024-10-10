# Smart_Excel_Extractor

![GitHub Profile](https://img.shields.io/badge/GitHub-mrkevler-green) ![Buy Me a Coffee](https://img.shields.io/badge/Support-Buy%20Me%20a%20Coffee-yellow) ![Python Version](https://img.shields.io/badge/Python-%3E%3D3.6-blue) ![License](https://img.shields.io/badge/License-CC%20BY--NC-blue)

![GitHub repo size](https://img.shields.io/github/repo-size/mrkevler/smart_excel_extractor) ![GitHub stars](https://img.shields.io/github/stars/mrkevler/smart_excel_extractor?style=social) ![GitHub forks](https://img.shields.io/github/forks/mrkevler/smart_excel_extractor?style=social) ![GitHub last commit](https://img.shields.io/github/last-commit/mrkevler/smart_excel_extractor)

`Extract categorized data from Excel files to CSV and TXT using a Python script. Easily convert your XLSX document into an import file with three fields: category, item and description.`

## Overview

Smart_Excel_Extractor is a Python script designed to intelligently extract data from Excel files (.xlsx) and output the data into both CSV and TXT formats. The script is prepared for structured documents that contain categories, actions and prompts. Additionally it handles merged cells gracefully. It helps users convert formatted prompt libraries into easy-to-use formats for further processing.

## Features

- Extracts fields from Excel (.xlsx) files.
- Handles merged cells and repeated rows intelligently.
- Filters out unnecessary data based on user-defined criteria.
- Outputs results in CSV and TXT formats.
- CLI interface to choose the Excel file.

## Technologies Used

- **Python**: The core programming language used for script logic and data processing.
- **openpyxl**: A library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files used for data extraction.
- **curses**: A terminal handling library used for creating interactive file selection.

## Python Version

This project requires **Python 3.6 or higher**.

## How to Use the Script

1. Clone the repository:

```sh
git clone https://github.com/mrkevler/smart_excel_extractor.git
```

2. Navigate to the directory:

```sh
cd smart_excel_extractor
```

3. Install the required libraries:

```sh
pip install -r requirements.txt
```

4. Run the script:

```sh
python smart_excel_extractor.py
```

The script will prompt you to select the Excel file from the current directory.

## Example

Given an Excel sheet with merged cells for categories and corresponding fields, the script will generate a CSV and TXT file with complete categorized data.

## Installation

```CLI
1. Clone the repository.
2. Make sure Python is installed.
3. Install the dependencies using the `requirements.txt` file:
   
   pip install -r requirements.txt
   
4. Run the script following the documentation.
```

## Support My Work

If you find this project helpful, consider supporting me on [Buy Me a Coffee](https://buymeacoffee.com/mrkevler).

## License

This project is licensed under **Creative Commons Attribution-NonCommercial (CC BY-NC)**. You are free to use and modify the code for non-commercial purposes, but you must give appropriate credit. For commercial use, please contact me to purchase attribution.

## Contact

Created by [mrkevler](https://github.com/mrkevler). You can contact me via [Telegram](https://t.me/mrkevler).

---

mrkevler
