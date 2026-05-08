import pdfplumber
import pandas as pd
import os
import logging
import argparse

# -------------------------------
# ARGUMENT PARSER (TERMINAL INPUT)
# -------------------------------
parser = argparse.ArgumentParser(description="PDF to Excel Extractor")

parser.add_argument("--pdf", required=True, help="Path to input PDF file")
parser.add_argument("--excel", required=True, help="Path to output Excel file")

args = parser.parse_args()

pdf_path = args.pdf
excel_path = args.excel

# -------------------------------
# LOGGING SETUP
# -------------------------------
logging.basicConfig(
    filename="pdf_extraction.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def log_info(message):
    print(message)
    logging.info(message)

def log_warning(message):
    print(f"⚠ {message}")
    logging.warning(message)

def log_error(message):
    print(f"❌ {message}")
    logging.error(message)

# -------------------------------
# KEYWORDS
# -------------------------------
FIELD_KEYWORDS = {
    "Tag Number": ["tag number", "tag no"],
    "Service": ["service"],
    "Line Size": ["line size"],
    "Fluid": ["fluid"],
    "Operating Pressure": ["operating pressure"],
    "Operating Temperature": ["operating temperature"],
    "Design Temperature": ["design temperature"],
    "Max Flow": ["maximum flow"],
    "Normal Flow": ["normal flow"],
    "Min Flow": ["minimum flow"],
    "Pipe ID": ["id"],
    "Pipe OD": ["od"],
    "Beta Ratio": ["beta"],
    "Bore Diameter": ["bore diameter"],
}

# -------------------------------
# EXTRACT TABLE DATA
# -------------------------------
def extract_tables(pdf_path):
    key_value_pairs = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            log_info(f"Opened PDF with {len(pdf.pages)} pages")

            for page_num, page in enumerate(pdf.pages, start=1):
                try:
                    tables = page.extract_tables()

                    for table in tables:
                        for row in table:
                            if not row:
                                continue

                            row = [str(cell).strip() if cell else "" for cell in row]

                            if len(row) >= 2:
                                key = row[0]
                                value = row[-1]

                                if key and value:
                                    key_value_pairs.append((key.lower(), value))

                except Exception as e:
                    log_error(f"Error processing page {page_num}: {str(e)}")

    except Exception as e:
        log_error(f"Failed to open PDF: {str(e)}")

    return key_value_pairs

# -------------------------------
# MATCH FIELDS
# -------------------------------
def map_fields(pairs, field_keywords):
    data = {}

    for field, keywords in field_keywords.items():
        found_value = "N/A"

        for key, value in pairs:
            for keyword in keywords:
                if keyword in key:
                    found_value = value
                    break
            if found_value != "N/A":
                break

        if found_value == "N/A":
            log_warning(f"{field} not found")

        data[field] = found_value

    return data

# -------------------------------
# SAVE TO EXCEL
# -------------------------------
def save_to_excel(data, excel_path):
    try:
        df_new = pd.DataFrame([data])

        if os.path.exists(excel_path):
            df_existing = pd.read_excel(excel_path)
            df_final = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_final = df_new

        df_final.to_excel(excel_path, index=False)
        log_info(f"Saved to {excel_path}")

    except Exception as e:
        log_error(f"Error saving Excel: {str(e)}")

# -------------------------------
# MAIN
# -------------------------------
def main():
    log_info("Starting PDF extraction...")

    if not os.path.exists(pdf_path):
        log_error("PDF file does not exist!")
        return

    pairs = extract_tables(pdf_path)

    if not pairs:
        log_warning("No structured data found in PDF.")
        return

    data = map_fields(pairs, FIELD_KEYWORDS)

    if data.get("Tag Number") == "N/A":
        log_warning("Tag Number not found!")

    save_to_excel(data, excel_path)

    log_info("Extraction complete.")
    print(data)


if __name__ == "__main__":
    main()


# python script.py --pdf "input.pdf" --excel "output.xlsx"
# venv\Scripts\Activate.ps1 ,activate virtual environmnet
# 