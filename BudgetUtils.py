import os
import re
import pandas as pd
import platform
import subprocess

# Set file paths
DATA_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Data")
CATEGORY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Categories.csv")
EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MonthlySpending.xlsm")


def load_categories():
    """Load categories from a CSV file if it exists."""
    if os.path.exists(CATEGORY_FILE):
        return pd.read_csv(CATEGORY_FILE, index_col=0).squeeze("columns").to_dict()
    return {} # Return empty dict if categories.csv DNE

categories = load_categories()


def save_categories(cats = categories):
    """Save categories dictionary to a CSV file, sorted alphabetically by category."""
    df = pd.DataFrame(list(cats.items()), columns=["Name", "Category"])
    df.sort_values(by=['Category', 'Name'], inplace=True)
    df.to_csv(CATEGORY_FILE, index=False)


def clean_text(text):
    """Removes prefixes, extra spaces, store codes, phone numbers, trailing numbers, and '.com' from business names."""
    text = text.upper()

    # Remove prefixes (case-insensitive, handles space between SQ and *)
    text = re.sub(r"^(TST\* ?|SQ ?\* ?)", "", text, flags=re.IGNORECASE)

    # Remove store codes (e.g., "#791", "# 20816")
    text = re.sub(r"#\s?\d+", "", text)

    # Remove phone numbers (e.g., "919-678-1444", "3122422019", "191-99518925")
    text = re.sub(r"\b\d{3,}-?\d{3,}-?\d{3,}\b", "", text)

    # Remove any occurrence of three or more digits not part of a word
    text = re.sub(r'\b\d{3,}\b', '', text)

    # Remove ".com"
    text = re.sub(r"\.com\b", "", text, flags=re.IGNORECASE)

    # Remove extra spaces and strip
    text = re.sub(r'\s+', ' ', text).strip()

    return text.upper()


def name_transaction(name):
    """Name a transaction, prompting user if unknown, and standardizing name."""
    standardized_name = name.upper()
    for key in categories:
        if key.lower() in standardized_name.lower():
            return key  # Return key if found

    # Ask user for name if unrecognized
    new_name = clean_text(input(f"Edit name for: [{name}] ") or name)
    categories[new_name] = "Uncategorized"
    
    save_categories()
    return new_name  # Return edited, standardized name


def categorize_transaction(name):
    """Categorize a transaction, prompting the user if unknown."""
    if categories.get(name) == "Uncategorized":  # Check if the specific name needs categorization
        category = input(f"Enter category for: [{name}] ").strip().lower().capitalize()
        categories[name] = category
        save_categories()
        return category
    return categories.get(name) # Return category if already exists


# List of known two-word city names
two_word_cities = {"CHAPEL HILL", 
                   "WINSTON SALEM", 
                   "WINSTON-SALEM", 
                   "SURF CITY", 
                   "MYRTLE BEACH", 
                   "NEW YORK", 
                   "SAN FRANCISCO", 
                   "LOS ANGELES"}

# Pattern to match phone numbers and website-like text (case-insensitive)
phone_or_web_pattern = re.compile(r'\b\d{3,}-?\d{3,}-?\d{3,}\b|https?://\S+|www\.\S+|[a-z0-9.-]+\.[a-z]{2,}', re.IGNORECASE)

def split_transaction(record):
    record = record.upper().strip()  # Convert everything to uppercase
    
    parts = record.split()
    
    # Extract the last two uppercase letters as the state if valid
    state = parts[-1] if re.match(r'^[A-Z]{2}$', parts[-1]) else ''
    
    # Identify city name
    city = ''
    if state and len(parts) > 1:
        # Check for a possible two-word city
        if len(parts) > 3 and " ".join(parts[-3:-1]) in two_word_cities:
            city = " ".join(parts[-3:-1])
        else:
            possible_city = parts[-2]
            # Ensure the city name is not a phone number or website
            if possible_city.isalpha() and not phone_or_web_pattern.search(possible_city):
                city = possible_city

    # The remaining part is the store name
    num_city_words = len(city.split()) if city else 0
    store_name = " ".join(parts[:-num_city_words - (1 if state else 0)])

    return pd.Series([store_name.strip(), city.strip(), state.strip()])


def file_input():
    # Create a list to hold (csv_file, sheet_name) tuples
    files_and_sheets = []

    while True:
        input_csv = input("\n\U0001F4C1 Input CSV File Name (FileName.csv or 'all') or press Enter to finish: ").strip().lower()

        if input_csv == "":
            break

        if input_csv == "all":
            data_folder = DATA_FOLDER
            try:
                csv_files = [f for f in os.listdir(data_folder) if f.endswith("Exp.csv")]
                if not csv_files:
                    print(f"\U0000274C No 'Exp.csv' files found in {data_folder}")
                for csv_file in csv_files:
                    csv_path = os.path.join(data_folder, csv_file)
                    sheet_name = csv_file[:-7]  # Remove "Exp.csv"
                    files_and_sheets.append((csv_path, sheet_name))
                break  # Exit the loop after processing "all"
            except FileNotFoundError:
                print(f"\U0000274C Data folder '{data_folder}' not found.")
                continue  # Go back to input prompt
        else:
            input_csv_path = os.path.join(DATA_FOLDER, input_csv)  # Full path to CSV
            default_sheet_name = (os.path.basename(input_csv)[:-7] if input_csv.endswith("exp.csv") else os.path.splitext(os.path.basename(input_csv))[0]).title()
            
            sheet_name = input("\U0001F4F0 New or Existing Excel Sheet Name (case-sensitive, press Enter to use default): ").strip()
            if sheet_name == "":
                sheet_name = default_sheet_name  # Use CSV filename without "Exp.csv" or extension

            files_and_sheets.append((input_csv_path, sheet_name))

    return files_and_sheets


def print_confirmation(files_and_sheets):
    # Confirmation message with CSV file names
    csv_names = [os.path.basename(csv_file) for csv_file, _ in files_and_sheets]
    sheet_names = [sheet for _, sheet in files_and_sheets]

    # Calculate the width for the table columns
    csv_width = max(len(name) for name in csv_names) + 1
    sheet_width = max(len(name) for name in sheet_names) + 1

    # Print the header
    print("\nProcessing transactions for:")
    print("-" * (csv_width + sheet_width + 7))
    print(f"| {'CSV Name'.ljust(csv_width)} | {'Sheet'.ljust(sheet_width)} |")
    print("-" * (csv_width + sheet_width + 7))

    # Print each row of CSV and sheet names
    for csv_name, sheet_name in zip(csv_names, sheet_names):
        print(f"| {csv_name.ljust(csv_width)} | {sheet_name.ljust(sheet_width)} |")

    print("-" * (csv_width + sheet_width + 7))


def open_excel_file():
    try:
        if platform.system() == 'Windows':
            os.startfile(EXCEL_FILE)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.call(['open', EXCEL_FILE])
        else:  # Linux (and others)
            subprocess.call(['xdg-open', EXCEL_FILE])  # Or a suitable command
    except Exception as e:
        print(f"\n\U0000274C Error opening file: {e}")