import pandas as pd
import os
import re

# Defining function to extract month from the "Last updated" text
def extract_month_from_text(text):
    month_map = {
        'january': '01', 'jan': '01', 'february': '02', 'feb': '02', 'march': '03', 'mar': '03',
        'april': '04', 'apr': '04', 'may': '05', 'june': '06', 'jun': '06', 'july': '07', 'jul': '07',
        'august': '08', 'aug': '08', 'september': '09', 'sep': '09', 'october': '10', 'oct': '10',
        'november': '11', 'nov': '11', 'december': '12', 'dec': '12'
    }
    # Extracting month name from text like "Last updated: September 2019"
    if isinstance(text, str):
        match = re.search(r'(january|february|march|april|may|june|july|august|september|october|november|december)', text.lower(), re.IGNORECASE)
        if match:
            month_name = match.group(1).lower()
            return month_map.get(month_name, 'Unknown')
    print(f"Warning: Could not extract month from text: {text}")
    return 'Unknown'

# Directory containing the Excel files
input_dir = r'D:\Users\paulli1\Desktop\MINI BUS Research\Market Data\2025'
output_file = 'processed_light_bus_data_2025_all_months.xlsx'

# List of expected filenames for Jan, Feb, and Mar 2025
expected_files = [
    'particulars of first registered vehicle_Jan 2025.xlsx',
    'particulars of first registered vehicle_Feb 2025.xlsx',
    'particulars of first registered vehicle_Mar 2025.xlsx'
]

# Listing all files in the directory for debugging
print(f"Files in directory {input_dir}:")
files_in_dir = os.listdir(input_dir)
for file in files_in_dir:
    print(f" - {file}")

# Initializing an empty list to store DataFrames
all_dfs = []

# Processing each expected file
for filename in expected_files:
    file_path = os.path.join(input_dir, filename)
    if not os.path.exists(file_path):
        print(f"Warning: File not found: {filename}. Skipping.")
        continue

    print(f"\nProcessing file: {filename}")

    # Reading the second row to get the "Last updated" text
    second_row = pd.read_excel(file_path, header=None, skiprows=1, nrows=1)
    last_updated_text = second_row.iloc[0, 0]  # Assuming the text is in the first column
    month = extract_month_from_text(last_updated_text)
    print(f"Extracted month: {month}")

    # Reading the main data with headers starting from the fourth row (0-based index 3)
    df = pd.read_excel(file_path, header=3)

    # Stripping whitespace from column names
    df.columns = df.columns.str.strip()

    # Printing column names for debugging
    print("Column names in the Excel file:", df.columns.tolist())

    # Printing first few rows to inspect data
    print("First few rows of the DataFrame:")
    print(df.head())

    # Finding the 'Vehicle Class' column (case-insensitive)
    vehicle_class_col = None
    for col in df.columns:
        if col.lower() == 'vehicle class':
            vehicle_class_col = col
            break

    if vehicle_class_col is None:
        print(f"Error: Column 'Vehicle Class' not found in {filename}. Skipping this file.")
        continue

    # Filtering for Vehicle Class 'Private Light Bus' or 'Public Light Bus'
    filtered_df = df[df[vehicle_class_col].isin(['Private Light Bus', 'Public Light Bus'])]

    # Adding 'Month' and 'Registration Year' columns
    filtered_df['Month'] = month
    filtered_df['Registration Year'] = 2025

    # Appending the filtered DataFrame to the list
    all_dfs.append(filtered_df)

# Combining all DataFrames into one
if all_dfs:
    combined_df = pd.concat(all_dfs, ignore_index=True)
    # Saving the combined data to a new Excel file
    combined_df.to_excel(output_file, index=False)
    print(f"\nProcessed data saved to {output_file}")
else:
    print("\nNo data was processed. Please check the input files and error messages above.")