# 🧹 Excel/CSV Data Cleaner & Report Automator

A Python tool that automatically cleans messy CSV/Excel data and generates a quick summary report — a task that comes up constantly in real business and data analysis workflows

![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python&logoColor=white)
![Status](https://img.shields.io/badge/status-active-brightgreen)
![License](https://img.shields.io/badge/license-MIT-lightgrey)

---

## 🎯 Problem It Solves

Raw data exports (from forms, sales systems, surveys, etc.) are almost never clean. Common issues include:

- Extra spaces in column names and text values
- Inconsistent capitalization (`"mouse"` vs `"MOUSE"` vs `"Mouse"`)
- Missing values (empty quantity, price, or customer name fields)
- Duplicate rows
- Inconsistent data types

This script **automates the cleanup process** instead of doing it manually in Excel, and produces a summary report showing what was cleaned.

---

## ✨ Features

- ✅ Loads data from `.csv` or `.xlsx` files
- ✅ Cleans column names and text formatting
- ✅ Fills missing values using sensible defaults (e.g., average price per product)
- ✅ Removes duplicate rows
- ✅ Generates a console report: total sales per product, total revenue, etc.
- ✅ Saves a cleaned version of the dataset as a new CSV file
- ✅ Easy to customize for your data structure

---

## 📊 Tech Stack

| Purpose | Library |
|---|---|
| Data manipulation | `pandas` |
| Excel files | `openpyxl` |

---

## 📂 Project Structure

```
csv-excel-data-cleaner/
├── data_cleaner.py           # Main script
├── data/
│   ├── raw_sales_data.csv    # Sample messy input data
│   └── cleaned_sales_data.csv # Generated after running
├── requirements.txt
└── README.md
```

---

## 🚀 Getting Started

### 1. Clone the repository
```bash
git clone https://github.com/ayeshamumtaz1057/csv-excel-data-cleaner.git
cd csv-excel-data-cleaner
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the script
```bash
python data_cleaner.py
```

### 4. Check the results

**Console output:** Summary report of cleaning operations  
**Generated file:** `data/cleaned_sales_data.csv` — your cleaned dataset

---

## 📋 Example Output

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

---

## 🔧 How to Customize for Your Data

1. **Replace sample data:** Replace `data/raw_sales_data.csv` with your own CSV/Excel file
2. **Update column names:** Edit the column names in `data_cleaner.py` if needed
3. **Adjust cleaning logic:** Modify the functions to match your specific data structure
4. **Change output:** Customize the report format in the console output section

---

## 📚 Key Concepts Demonstrated

- Using `pandas` for real-world data cleaning
- Handling missing data with group-based logic
- Structuring a Python script into clear, reusable functions
- Writing a project others can actually clone and run
- Building a practical tool for business data

---

## 💡 What I Learned Building This

- How to work with `pandas` DataFrames efficiently
- Strategies for handling missing data intelligently
- Creating reusable, maintainable functions
- Generating meaningful reports from raw data
- Building tools that solve real problems

---

## 🚀 Possible Future Improvements

- [ ] Add support for cleaning multiple files in batch
- [ ] Export the report as PDF or Excel summary sheet
- [ ] Add data visualization (charts) using matplotlib
- [ ] Add command-line arguments for custom file paths
- [ ] Implement logging for detailed cleaning operations
- [ ] Add data validation checks
- [ ] Support for different file formats (JSON, XML)

---

## 💼 Real-World Use Cases

- **E-commerce:** Clean product inventory and sales data
- **HR Analytics:** Process employee records and surveys
- **Finance:** Prepare financial reports from raw exports
- **Marketing:** Clean customer databases and contact lists
- **Research:** Prepare datasets for analysis

---

## 👩‍💻 Author

**Ayesha Mumtaz** — BS Information Technology Student  
📍 Bahawalpur, Punjab, Pakistan  
🔗 [LinkedIn](https://www.linkedin.com/in/ayesha-mumtaz-82b8913a9)  
🔗 [GitHub](https://github.com/ayeshamumtaz1057)

---

## 📄 License

This project is open-sourced for educational purposes under the MIT License.

---

## ⭐ Contribute or Suggest

Have ideas to improve this tool? Feel free to:
- ⭐ Star this repo
- 🍴 Fork and submit a pull request
- 💬 Open an issue with suggestions
