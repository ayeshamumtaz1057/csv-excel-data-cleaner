# Excel/CSV Data Cleaner & Report Automator

A Python tool that automatically cleans messy CSV/Excel data and generates
a quick summary report — a task that comes up constantly in real business
and data analysis workflows.

## Problem it solves

Raw data exports (from forms, sales systems, surveys, etc.) are almost
never clean. Common issues include:
- Extra spaces in column names and text values
- Inconsistent capitalization (`"mouse"` vs `"MOUSE"` vs `"Mouse"`)
- Missing values (empty quantity, price, or customer name fields)
- Duplicate rows

This script automates the cleanup process instead of doing it manually
in Excel, and produces a short report summarizing the results.

## Features

- Loads data from `.csv` or `.xlsx` files
- Cleans column names and text formatting
- Fills missing values using sensible defaults (e.g. average price per product)
- Removes duplicate rows
- Generates a console report: total sales per product, total revenue, etc.
- Saves a cleaned version of the dataset as a new CSV file

## Project structure

```
python-scraping-automation/
├── data_cleaner.py          # Main script
├── data/
│   ├── raw_sales_data.csv       # Sample messy input data
│   └── cleaned_sales_data.csv   # Generated after running the script
├── requirements.txt
└── README.md
```

## How to run

1. Clone the repository:
   ```
   git clone https://github.com/ayeshamumtaz1057/python-scraping-automation.git
   cd python-scraping-automation
   ```

2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Run the script:
   ```
   python data_cleaner.py
   ```

4. Check the console for the summary report, and find the cleaned file at
   `data/cleaned_sales_data.csv`.

## Example output

```
========================================
        DATA CLEANING REPORT
========================================
Total rows after cleaning : 10
Duplicate rows removed    : 0

Total quantity sold per product:
Product
Keyboard    4.0
Laptop      7.0
Monitor     1.0
Mouse       9.0
Webcam      2.0

Total revenue (approx): 552,200
========================================
```

## What I learned building this

- Using `pandas` for real-world data cleaning (not just tutorials)
- Handling missing data with group-based logic instead of blanket fixes
- Structuring a Python script into clear, reusable functions
- Writing a project others can actually clone and run

## Possible future improvements

- Add support for cleaning multiple files in a batch
- Export the report as a PDF or Excel summary sheet
- Add data visualization (charts) using matplotlib
- Add command-line arguments to specify input/output file paths

## Author

**Ayesha Mumtaz**
BS Information Technology student | Aspiring AI/ML Engineer
[LinkedIn](https://www.linkedin.com/in/ayesha-mumtaz-82b8913a9)
