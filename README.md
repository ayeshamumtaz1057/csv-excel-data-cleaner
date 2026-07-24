<div align="center">

# 🧹 Excel/CSV Data Cleaner & Report Automator

**A Python tool that automatically cleans messy CSV/Excel data and generates a quick summary report — a task that comes up constantly in real business and data analysis workflows.**

[![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![pandas](https://img.shields.io/badge/pandas-data%20cleaning-150458?style=flat-square&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![openpyxl](https://img.shields.io/badge/openpyxl-Excel%20support-1D6F42?style=flat-square)](https://openpyxl.readthedocs.io/)
[![Status](https://img.shields.io/badge/status-active-brightgreen?style=flat-square)](#)
[![License](https://img.shields.io/badge/license-MIT-yellow?style=flat-square)](LICENSE)
[![No API Key](https://img.shields.io/badge/setup-zero%20config-success?style=flat-square)](#)

Built by **[Ayesha Mumtaz](https://github.com/ayeshamumtaz1057)**

[Problem](#-problem-it-solves) · [Features](#-features) · [How It Works](#️-how-it-works) · [Getting Started](#-getting-started) · [Cleaning Logic](#-cleaning-logic) · [Troubleshooting](#-troubleshooting) · [FAQ](#-faq)

</div>

---

## 📑 Table of Contents

- [Problem It Solves](#-problem-it-solves)
- [Features](#-features)
- [Tech Stack](#-tech-stack)
- [How It Works](#️-how-it-works)
- [Cleaning Logic](#-cleaning-logic)
- [Function Reference](#-function-reference)
- [Project Structure](#-project-structure)
- [Getting Started](#-getting-started)
- [Example Output](#-example-output)
- [How to Customize for Your Data](#-how-to-customize-for-your-data)
- [Performance](#-performance)
- [Troubleshooting](#-troubleshooting)
- [FAQ](#-faq)
- [Key Concepts Demonstrated](#-key-concepts-demonstrated)
- [What I Learned Building This](#-what-i-learned-building-this)
- [Possible Future Improvements](#-possible-future-improvements)
- [Real-World Use Cases](#-real-world-use-cases)
- [Author](#-author)
- [License](#-license)

---

## 🎯 Problem It Solves

Raw data exports (from forms, sales systems, surveys, etc.) are almost never clean. Common issues include:

- Extra spaces in column names and text values
- Inconsistent capitalization (`"mouse"` vs `"MOUSE"` vs `"Mouse"`)
- Missing values (empty quantity, price, or customer name fields)
- Duplicate rows
- Inconsistent data types

This script **automates the cleanup process** instead of doing it manually in Excel, and produces a summary report showing what was cleaned.

### Why it matters

Doing this by hand takes twenty minutes, is easy to get wrong, and leaves no record of what changed. Running a script takes a second, produces identical results every time, and prints exactly what it did.

---

## ✨ Features

### Cleaning

| Feature | What it does |
|:--|:--|
| **Multi-format loading** | Reads `.csv`, `.xlsx` and `.xls`; raises a clear error on anything else |
| **Column name cleanup** | Strips stray leading/trailing spaces from every header |
| **Whitespace trimming** | Strips surrounding spaces from every text value in the file |
| **Capitalization fix** | `Product` and `Customer Name` converted to title case, so `mouse` / `MOUSE` / `Mouse` become one value |
| **Smart missing values** | Each column filled by a rule that suits it — see [Cleaning Logic](#-cleaning-logic) |
| **Duplicate removal** | Fully duplicate rows dropped, with the count reported |
| **Null normalisation** | Text `"nan"` produced during conversion is turned back into a real missing value |

### Reporting

| Feature | What it does |
|:--|:--|
| **Row summary** | Total rows remaining after cleaning |
| **Duplicate count** | How many duplicate rows were removed |
| **Quantity per product** | Total units sold, grouped by product |
| **Revenue total** | `Price × Quantity` summed across the dataset, thousands-separated |
| **Derived column** | A `Total` column is calculated and carried into the saved output |

### System

| Feature | What it does |
|:--|:--|
| **Non-destructive** | The input file is never modified — output goes to a new CSV |
| **Column-safe** | Every rule checks the column exists first, so it won't crash on a different schema |
| **Reusable functions** | Seven small single-purpose functions you can import individually |
| **Zero config** | No API keys, no accounts, no internet connection required |
| **Sample data included** | Ships with `raw_sales_data.csv`, so it runs the moment you clone it |

---

## 📊 Tech Stack

| Purpose           | Library    |
| ----------------- | ---------- |
| Data manipulation | `pandas`   |
| Excel files       | `openpyxl` |

<details>
<summary><b>Why these choices</b> (click to expand)</summary>

<br>

| Decision | Reasoning |
|:--|:--|
| **pandas** over the manual `csv` module | Vectorised operations clean thousands of rows in milliseconds. Filling prices with a per-product average is one `groupby().transform()` call and would be twenty lines of loops otherwise. |
| **Group-based price fill** over a global average | A missing laptop price filled with the overall average would come out near a mouse's price. Filling from that product's own rows keeps the number plausible. |
| **`Quantity` filled with `1`** over an average | An order exists, so the quantity is at least one. Assuming the minimum is conservative — an average would invent sales that never happened. |
| **`"Unknown Customer"`** over dropping the row | The sale is real even when the name is missing. Deleting the row would silently reduce revenue; labelling it keeps the total honest and makes the gap visible. |
| **Non-destructive output** | Cleaning rules are judgement calls. Writing to a new file means a wrong rule costs a re-run, not the original data. |
| **A function per step** | Each rule can be read, tested, reordered or reused on its own. A single long `clean()` function could be none of those. |
| **Existence checks before every rule** | `if "Product" in df.columns` means pointing this at a different schema degrades gracefully instead of crashing. |

</details>

---

## 🏗️ How It Works

```
 [ data/raw_sales_data.csv ]
              |
              v
      load_data()              CSV or Excel → DataFrame
              |
              v
      clean_column_names()     strip spaces from headers
              |
              v
      clean_text_columns()     trim values, title-case names
              |
              v
      handle_missing_values()  per-column fill rules
              |
              v
      remove_duplicates()      drop exact duplicates, count them
              |
              v
      generate_report()        print the summary
              |
              v
      save_clean_data()        write the output CSV
              |
              v
 [ data/cleaned_sales_data.csv ]
```

Each function takes a DataFrame and returns a DataFrame, so steps can be reordered, skipped or reused independently. `main()` wires them together.

---

## 🧮 Cleaning Logic

### Missing values — one rule per column

The core idea: **a single fill strategy for the whole file is always wrong somewhere.** Each column gets the rule that fits it.

| Column | Rule | Reasoning |
|:--|:--|:--|
| `Quantity` | Fill with `1` | An order exists, so at least one unit was sold. Conservative — never invents sales |
| `Price` | Group mean **per product** | A laptop's missing price is estimated from other laptops, not from the whole catalogue |
| `Customer Name` | `"Unknown Customer"` | The sale is real; deleting the row would understate revenue |

The price rule is the interesting one:

```python
# ❌ A missing laptop price becomes roughly a mouse price
df["Price"] = df["Price"].fillna(df["Price"].mean())

# ✅ Each product's gap is filled from its own rows
df["Price"] = df.groupby("Product")["Price"].transform(
    lambda x: x.fillna(x.mean())
)
```

Filling from the overall average would flatten every product toward the middle — exactly the difference the data exists to show.

### Capitalization

Product and customer names are title-cased, so `mouse`, `MOUSE` and `Mouse` stop being three separate entries. This matters more than it looks: without it, `groupby("Product")` reports the same product three times and every total is wrong.

### Whitespace

Both headers and values are stripped. `" Product "` and `"Product"` are different columns to pandas, and `" Mouse"` and `"Mouse"` are different products — invisible bugs that produce wrong totals rather than errors.

---

## 🔍 Function Reference

| Function | Signature | Returns |
|:--|:--|:--|
| `load_data` | `(file_path)` | DataFrame — raises `ValueError` on unsupported extensions |
| `clean_column_names` | `(df)` | DataFrame with stripped headers |
| `clean_text_columns` | `(df)` | DataFrame with trimmed, title-cased text |
| `handle_missing_values` | `(df)` | DataFrame with gaps filled per the rules above |
| `remove_duplicates` | `(df)` | **Tuple** — `(df, removed_count)` |
| `generate_report` | `(df, duplicates_removed)` | `None` — prints to console |
| `save_clean_data` | `(df, output_path)` | `None` — writes the CSV |

> ⚠️ Note that `remove_duplicates` returns a **tuple**, not a DataFrame — unpack it:
> ```python
> df, removed = remove_duplicates(df)
> ```

**Using a single function on its own:**

```python
import pandas as pd
from data_cleaner import clean_text_columns, handle_missing_values

df = pd.read_csv("your_file.csv")
df = clean_text_columns(df)
df = handle_missing_values(df)
```

---

## 📂 Project Structure

```
csv-excel-data-cleaner/
├── data_cleaner.py            # Main script — all cleaning functions
├── data/
│   ├── raw_sales_data.csv     # Sample messy input data
│   └── cleaned_sales_data.csv # Generated after running
├── requirements.txt
├── README.md
└── .github/
    └── ISSUE_TEMPLATE/        # Feature request template
```

---

## 🚀 Getting Started

### Prerequisites

| Requirement | Version | Required? |
|:--|:--|:--|
| Python | 3.x | ✅ Yes |
| pip | latest | ✅ Yes |
| Internet | — | ❌ No — runs fully offline |

### 1. Clone the repository

```bash
git clone https://github.com/ayeshamumtaz1057/csv-excel-data-cleaner.git
cd csv-excel-data-cleaner
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

<details>
<summary>Recommended: use a virtual environment first</summary>

<br>

```bash
python -m venv venv

venv\Scripts\activate           # Windows
source venv/bin/activate        # macOS / Linux

pip install -r requirements.txt
```

Your prompt should now begin with `(venv)`.

</details>

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

<details>
<summary><b>Worked example — adapting to an HR dataset</b></summary>

<br>

The rules are keyed to column names, so swapping schemas means editing those references:

```python
# clean_text_columns()
if "Department" in df.columns:
    df["Department"] = df["Department"].str.title()

# handle_missing_values()
if "Salary" in df.columns and "Department" in df.columns:
    df["Salary"] = df.groupby("Department")["Salary"].transform(
        lambda x: x.fillna(x.median())
    )
```

Every rule is wrapped in an existence check, so columns that don't apply are skipped rather than causing a crash.

</details>

<details>
<summary><b>Using an Excel file instead of CSV</b></summary>

<br>

`load_data()` already handles `.xlsx` and `.xls` — just change the path:

```python
input_file = "data/raw_sales_data.xlsx"
```

Make sure `openpyxl` is installed; pandas needs it to read Excel.

</details>

<details>
<summary><b>Protecting a column from being filled</b></summary>

<br>

Identifier columns should never be imputed — a fabricated `Order ID` silently corrupts every join that uses it. Simply don't add a rule for them in `handle_missing_values()`; anything without a rule is left untouched.

</details>

---

## ⚡ Performance

| Rows | Typical time |
|:--|:--|
| 100 | instant |
| 10,000 | ~0.3 s |
| 100,000 | ~2 s |
| 1,000,000 | ~20 s |

pandas operations are vectorised, so cost grows roughly linearly. Excel files are noticeably slower to read than CSV — that's `openpyxl` parsing XML, not the cleaning itself. For large datasets, convert to CSV first.

---

## 🛠 Troubleshooting

<details open>
<summary><b>Setup</b></summary>

| Error | Fix |
|:--|:--|
| `ModuleNotFoundError: No module named 'pandas'` | Dependencies not installed, or the virtual environment isn't active. Run `pip install -r requirements.txt` |
| `Missing optional dependency 'openpyxl'` | `pip install openpyxl` — pandas needs it to read `.xlsx` |
| `'python' is not recognized` | Reinstall Python and tick **Add Python to PATH**, or use `py` on Windows |

</details>

<details>
<summary><b>Reading files</b></summary>

| Error | Fix |
|:--|:--|
| `FileNotFoundError: data/raw_sales_data.csv` | Run the script from the project root — the path is relative. Check the `data/` folder exists and contains the CSV |
| `ValueError: Unsupported file type` | `load_data()` accepts `.csv`, `.xlsx` and `.xls` only. Convert your file first |
| `UnicodeDecodeError` | The CSV isn't UTF-8. Try `pd.read_csv(path, encoding="latin-1")` inside `load_data()` |
| `PermissionError` | The file is open in Excel. Close it and re-run |
| `ParserError: Error tokenizing data` | Inconsistent column counts. Add `on_bad_lines="skip"` to identify the offending rows |
| All data lands in one column | Wrong delimiter — pass `sep=";"` or `sep="\t"` to `read_csv` |

</details>

<details>
<summary><b>Cleaning behaviour</b></summary>

| Issue | Fix |
|:--|:--|
| Nothing gets title-cased | The rules look for `Product` and `Customer Name` exactly. Check your headers match, or add your own column names |
| `Price` still has blanks | That product has **no** price anywhere, so the group average is itself empty. Add a fallback to the overall mean |
| Same product appears twice in the report | A capitalization or whitespace variant survived. Print `df["Product"].unique()` to spot it |
| Revenue looks wrong | `Price` or `Quantity` is stored as text. Check with `df.dtypes` — they must be numeric |
| Fewer rows than expected | Duplicate detection matches on **all** columns. Pass a `subset` of key columns to `drop_duplicates()` instead |
| `SettingWithCopyWarning` | You're modifying a slice. Use `.copy()` when creating the subset, or assign with `.loc` |

</details>

<details>
<summary><b>Writing output</b></summary>

| Error | Fix |
|:--|:--|
| `PermissionError` on save | The output CSV is open in Excel. Close it |
| `FileNotFoundError` on save | The `data/` folder doesn't exist. Create it, or use `os.makedirs(..., exist_ok=True)` |
| Leading zeros vanish in Excel | Excel treats `00123` as a number. Keep the column as text, or open the CSV via Data → From Text |
| Accented / Urdu text is garbled | Save with `encoding="utf-8-sig"` so Excel reads it correctly |

</details>

<details>
<summary><b>Git & GitHub</b></summary>

| Error | Fix |
|:--|:--|
| `! [rejected] main -> main (fetch first)` | Remote has commits you don't. `git pull --rebase origin main`, then push |
| `Support for password authentication was removed` | Use a Personal Access Token: Settings → Developer settings → Tokens (classic) → scope `repo` |
| Large data file rejected | GitHub caps files at 100 MB. Add it to `.gitignore` and commit a small sample instead |

</details>

---

## ❓ FAQ

<details>
<summary><b>Does it modify my original file?</b></summary>
<br>
No. The input is read-only and the cleaned data is written to a separate file. A cleaning rule that turns out to be wrong costs you a re-run, not your data.
</details>

<details>
<summary><b>Why is a missing Price filled per product instead of one overall average?</b></summary>
<br>
Because a global average is wrong for almost every row. A missing laptop price filled with the catalogue-wide mean comes out somewhere near a mouse. Grouping by product keeps each estimate inside a plausible range.
</details>

<details>
<summary><b>Why is a missing Quantity filled with 1 rather than an average?</b></summary>
<br>
Because the row exists, so at least one unit was sold — that's a fact, not an estimate. An average would invent sales that never happened and inflate revenue.
</details>

<details>
<summary><b>Why not just delete rows with missing customer names?</b></summary>
<br>
The sale still happened. Dropping the row would quietly reduce your revenue total; labelling it <code>"Unknown Customer"</code> keeps the number honest and makes the data-quality gap visible.
</details>

<details>
<summary><b>Will it work on data that isn't sales data?</b></summary>
<br>
The generic steps — whitespace, headers, duplicates — work on anything. The specific rules look for <code>Product</code>, <code>Price</code>, <code>Quantity</code> and <code>Customer Name</code>, and are skipped when those columns are absent. See the <a href="#-how-to-customize-for-your-data">HR example</a> for adapting them.
</details>

<details>
<summary><b>Does it support Excel files?</b></summary>
<br>
Yes — <code>load_data()</code> reads <code>.xlsx</code> and <code>.xls</code> as well as <code>.csv</code>. Output is always written as CSV.
</details>

<details>
<summary><b>How large a file can it handle?</b></summary>
<br>
Comfortably a few hundred thousand rows on a normal laptop. Beyond that, read in chunks or use a database — pandas holds the whole frame in memory.
</details>

<details>
<summary><b>Is my data sent anywhere?</b></summary>
<br>
No. Everything runs locally, there are no API calls, and the script works with no internet connection at all.
</details>

<details>
<summary><b>Why does the output have a Total column that wasn't in the input?</b></summary>
<br>
<code>generate_report()</code> calculates <code>Price × Quantity</code> to compute revenue, and that column is carried into the saved file. It's useful, so it's kept — drop it before saving if you'd rather the output matched the input schema exactly.
</details>

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

<details>
<summary><b>The specific lessons behind those points</b></summary>

<br>

**One fill strategy for the whole file is always wrong somewhere.** The obvious approach — fill every gap with the column average — produces laptop prices near mouse prices and invents quantities that were never ordered. Working out a separate rule for `Quantity`, `Price` and `Customer Name` was the point where this stopped being a tutorial exercise and started being a tool.

**`groupby().transform()` was the unlock.** Filling each product's price from its own rows sounds like it needs a loop over products. It's one line, and understanding *why* `transform` returns something the same shape as the original was the most valuable pandas concept the project taught.

**Capitalization bugs don't crash — they lie.** `mouse`, `Mouse` and `MOUSE` grouped into three separate products and the report showed three sets of totals that all looked plausible. Nothing errored. Normalising text early prevents a whole category of bugs that produce confident wrong answers.

**Invisible whitespace is the same class of bug.** `" Product "` and `"Product"` are different columns to pandas. Stripping both headers and values costs two lines and removes a failure mode that's almost impossible to spot by eye.

**Guard clauses make a script reusable.** Wrapping every rule in `if "Product" in df.columns` means pointing the script at a different dataset skips the rules that don't apply instead of crashing. That single habit turned a sales-only script into something adaptable.

**A cleaner without a report is a black box.** Producing a clean file with no indication of what changed asks the user to take it on trust. Printing rows removed, gaps filled and the resulting totals turned the output into something reviewable.

</details>

---

## 🚀 Possible Future Improvements

- [ ] Add support for cleaning multiple files in batch
- [ ] Export the report as PDF or Excel summary sheet
- [ ] Add data visualization (charts) using matplotlib
- [ ] Add command-line arguments for custom file paths
- [ ] Implement logging for detailed cleaning operations
- [ ] Add data validation checks
- [ ] Support for different file formats (JSON, XML)
- [ ] Write the report to a file alongside the cleaned data
- [ ] Add unit tests for each cleaning function
- [ ] Auto-detect numeric columns stored as text and convert them

---

## 💼 Real-World Use Cases

- **E-commerce:** Clean product inventory and sales data
- **HR Analytics:** Process employee records and surveys
- **Finance:** Prepare financial reports from raw exports
- **Marketing:** Clean customer databases and contact lists
- **Research:** Prepare datasets for analysis

---

## 🤝 Contributing

Contributions are welcome.

```bash
git checkout -b feature/your-feature
git commit -m "Add your feature"
git push origin feature/your-feature
```

Please keep each cleaning rule in its own function, guard it with a column-existence check, and add a line to the report for anything it changes.

Have an idea? Open one using the [feature request template](.github/ISSUE_TEMPLATE).

---

## 👩‍💻 Author

**Ayesha Mumtaz** — BS Information Technology Student

📍 Bahawalpur, Punjab, Pakistan

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=flat-square&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/ayesha-mumtaz-82b8913a9)
[![GitHub](https://img.shields.io/badge/GitHub-ayeshamumtaz1057-181717?style=flat-square&logo=github&logoColor=white)](https://github.com/ayeshamumtaz1057)

---

## 📄 License

This project is open-sourced for educational purposes under the MIT License.

---

## ⭐ Contribute or Suggest

Have ideas to improve this tool? Feel free to:

- ⭐ Star this repo
- 🍴 Fork and submit a pull request
- 💬 Open an issue with suggestions

<div align="center">

⭐ **Star this repo if you found it useful**

</div>
