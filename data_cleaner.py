"""
Excel/CSV Data Cleaner & Report Automator
-------------------------------------------
Author: Ayesha Mumtaz

What this script does:
1. Reads a messy CSV/Excel file (sales data in this example)
2. Cleans it:
   - Removes extra spaces from column names and text values
   - Fixes inconsistent capitalization (e.g. "mouse" vs "MOUSE")
   - Removes duplicate rows
   - Fills or flags missing values
3. Generates a clean output file
4. Prints a short summary report (like a mini business report)

This is meant to simulate a real task: businesses often receive messy
data exports and need it cleaned before analysis or reporting.
"""

import pandas as pd
import os


def load_data(file_path):
    """Load data from a CSV or Excel file into a pandas DataFrame."""
    if file_path.endswith(".csv"):
        df = pd.read_csv(file_path)
    elif file_path.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file type. Use .csv or .xlsx")
    return df


def clean_column_names(df):
    """Remove extra spaces and standardize column names."""
    df.columns = [col.strip() for col in df.columns]
    return df


def clean_text_columns(df):
    """Strip whitespace and standardize capitalization in text columns."""
    text_columns = df.select_dtypes(include=["object", "str"]).columns

    for col in text_columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace("nan", pd.NA)

    # Standardize product names and customer names to title case
    if "Product" in df.columns:
        df["Product"] = df["Product"].str.title()

    if "Customer Name" in df.columns:
        df["Customer Name"] = df["Customer Name"].str.title()

    return df


def handle_missing_values(df):
    """
    Handle missing values:
    - Missing Quantity -> fill with 1 (assume minimum order)
    - Missing Price -> fill with the average price of that product
    - Missing Customer Name -> mark as 'Unknown Customer'
    """
    if "Quantity" in df.columns:
        df["Quantity"] = df["Quantity"].fillna(1)

    if "Price" in df.columns and "Product" in df.columns:
        df["Price"] = df.groupby("Product")["Price"].transform(
            lambda x: x.fillna(x.mean())
        )

    if "Customer Name" in df.columns:
        df["Customer Name"] = df["Customer Name"].fillna("Unknown Customer")

    return df


def remove_duplicates(df):
    """Remove fully duplicate rows and return the cleaned DataFrame + count removed."""
    before = len(df)
    df = df.drop_duplicates()
    after = len(df)
    removed = before - after
    return df, removed


def generate_report(df, duplicates_removed):
    """Print a short summary report of the cleaned data."""
    print("\n" + "=" * 40)
    print("        DATA CLEANING REPORT")
    print("=" * 40)
    print(f"Total rows after cleaning : {len(df)}")
    print(f"Duplicate rows removed    : {duplicates_removed}")

    if "Product" in df.columns and "Quantity" in df.columns:
        print("\nTotal quantity sold per product:")
        print(df.groupby("Product")["Quantity"].sum().to_string())

    if "Price" in df.columns and "Quantity" in df.columns:
        df["Total"] = df["Price"] * df["Quantity"]
        total_revenue = df["Total"].sum()
        print(f"\nTotal revenue (approx): {total_revenue:,.0f}")

    print("=" * 40 + "\n")


def save_clean_data(df, output_path):
    """Save the cleaned DataFrame to a new CSV file."""
    df.to_csv(output_path, index=False)
    print(f"Clean file saved to: {output_path}")


def main():
    input_file = "data/raw_sales_data.csv"
    output_file = "data/cleaned_sales_data.csv"

    print(f"Loading data from {input_file} ...")
    df = load_data(input_file)

    df = clean_column_names(df)
    df = clean_text_columns(df)
    df = handle_missing_values(df)
    df, duplicates_removed = remove_duplicates(df)

    generate_report(df, duplicates_removed)
    save_clean_data(df, output_file)


if __name__ == "__main__":
    main()
