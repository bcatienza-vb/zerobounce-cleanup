import sys
import os
import pandas as pd
from openpyxl import load_workbook

def clean_csv_to_xlsx(input_csv, output_xlsx=None):
    # Try reading CSV with utf-8, fallback to latin1 if fails
    try:
        df = pd.read_csv(input_csv, encoding="utf-8", dtype=str)
    except UnicodeDecodeError:
        df = pd.read_csv(input_csv, encoding="latin1", dtype=str)

    # Ensure all columns are string
    df.columns = df.columns.astype(str)

    # Define mapping for columns
    keepCols = {
        "First Name": "First Name",
        "Last Name": "Last Name",
        "Email": "Email",
        "Phone Number": "Phone Number",
        "Mobile": "Phone Number",
        "Job Title": "Job Title",
        "Company Name": "Company",
        "Industries": "Industry",
        "Company Post Code/ZIP": "Postcode",
        "Industry": "Industry",
        "Country": "Country",
        "City": "City",
        "Postcode": "Postcode",
        "Website URL": "Website URL",
        "Website": "Website URL",
        "Marketing Consent": "Marketing Consent",
        "Lead Source": "Lead Source"
    }

    # Keep only relevant columns
    df = df[[col for col in df.columns if col in keepCols]]

    # Rename columns
    df = df.rename(columns=keepCols)

    # Rearrange columns
    newOrder = [
        "First Name", "Last Name", "Email", "Phone Number",
        "Job Title", "Company", "Industry", "Country",
        "City", "Postcode", "Website URL",
        "Marketing Consent", "Lead Source"
    ]
    df = df[[col for col in newOrder if col in df.columns]]

    # Clean Phone Number column
    if "Phone Number" in df.columns:
        df["Phone Number"] = (
            df["Phone Number"]
            .astype(str)
            .str.replace("\t", "", regex=False)          # remove tabs
            .str.replace(r"(?!^)\+", "", regex=True)     # remove + unless at start
            .str.replace(r"[^\d+]", "", regex=True)      # keep only digits and leading +
        )
        # Add apostrophe in front of all phone numbers
        df["Phone Number"] = df["Phone Number"].apply(lambda x: f"'{x}" if x else x)

    # Set default output filename
    if not output_xlsx:
        base, _ = os.path.splitext(input_csv)
        output_xlsx = base + ".xlsx"

    # Save as Excel
    df.to_excel(output_xlsx, index=False)

    # Adjust column widths to 20
    wb = load_workbook(output_xlsx)
    ws = wb.active
    for col in ws.columns:
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 20
    wb.save(output_xlsx)

    print(f"\n✅ File cleaned, formatted, and saved as: {output_xlsx}\n")
    input("Press Enter to exit")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Drag and drop a CSV file onto this script, or run:\n")
        print("    python clean_csv.py yourfile.csv\n")
        input("Press Enter to exit")
    else:
        input_file = sys.argv[1]
        clean_csv_to_xlsx(input_file)
