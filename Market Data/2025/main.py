import pandas as pd
import os
import re
import numpy as np

# Set pandas option to avoid FutureWarning for downcasting
pd.set_option('future.no_silent_downcasting', True)

# Defining function to extract month from the "Last updated" text
def extract_month_from_text(text):
    month_map = {
        'january': '01', 'jan': '01', 'february': '02', 'feb': '02', 'march': '03', 'mar': '03',
        'april': '04', 'apr': '04', 'may': '05', 'june': '06', 'jun': '06', 'july': '07', 'jul': '07',
        'august': '08', 'aug': '08', 'september': '09', 'sep': '09', 'october': '10', 'oct': '10',
        'november': '11', 'nov': '11', 'december': '12', 'dec': '12'
    }
    # Extracting month name from text like "Last updated: September 2025"
    if isinstance(text, str):
        match = re.search(r'(january|february|march|april|may|june|july|august|september|october|november|december)', text.lower(), re.IGNORECASE)
        if match:
            month_name = match.group(1).lower()
            return month_map.get(month_name, 'Unknown')
    print(f"Warning: Could not extract month from text: {text}")
    return 'Unknown'

# Extracting month from filename
def extract_month_from_filename(filename):
    month_patterns = {
        r'jan': '01', r'feb': '02', r'mar': '03', r'apr': '04',
        r'may': '05', r'jun': '06', r'jul': '07', r'aug': '08',
        r'sep': '09', r'oct': '10', r'nov': '11', r'dec': '12'
    }
    
    filename_lower = filename.lower()
    for pattern, month_num in month_patterns.items():
        if pattern in filename_lower:
            return month_num
    
    print(f"Warning: Could not extract month from filename: {filename}")
    return 'Unknown'

# Directory containing the Excel files
input_dir = r'D:\Users\paulli1\Desktop\Project Flexis\Market Data\2025'
output_file = 'processed_LGV_data_2025_all_months.xlsx'

# List of expected filenames for Jan to Jun 2025 (based on current date: July 7, 2025)
expected_files = [
    'particulars of first registered vehicle_Jan 2025.xlsx',
    'particulars of first registered vehicle_Feb 2025.xlsx',
    'particulars of first registered vehicle_Mar 2025.xlsx',
    'particulars of first registered vehicle_Apr 2025.xlsx',
    'particulars of first registered vehicle_May 2025.xlsx',
    'particulars of first registered vehicle_Jun 2025.xlsx'
]

# Listing all files in the directory for debugging
print(f"Files in directory {input_dir}:")
files_in_dir = os.listdir(input_dir)
for file in files_in_dir:
    print(f" - {file}")

# Initializing an empty list to store DataFrames
all_dfs = []

# Required columns
required_columns = [
    'Vehicle Class', 'Vehicle Make', 'Vehicle Model', 'Fuel Type',
    'Cylinder Capacity Of Engine (c.c.)', 'Rated Power (kW)', 'Body Type',
    'First Registration Vehicle Status (Note)', 'Permitted Gross Vehicle Weight',
    'Number Of Passenger Seats', 'Taxable Value (HK$)', 'Year Of Manufacture'
]

# Processing each expected file
for filename in expected_files:
    file_path = os.path.join(input_dir, filename)
    if not os.path.exists(file_path):
        print(f"Warning: File not found: {filename}. Skipping.")
        continue

    print(f"\nProcessing file: {filename}")
    
    # Extracting month from filename
    month = extract_month_from_filename(filename)
    print(f"Extracted month from filename: {month}")

    try:
        # Attempt to read the second row to get the "Last updated" text
        try:
            second_row = pd.read_excel(file_path, header=None, skiprows=1, nrows=1)
            last_updated_text = second_row.iloc[0, 0]  # Assuming the text is in the first column
            month_from_text = extract_month_from_text(last_updated_text)
            if month_from_text != 'Unknown':
                month = month_from_text
                print(f"Extracted month from 'Last updated' text: {month}")
        except Exception as e:
            print(f"Error reading 'Last updated' text: {e}, using month from filename")

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
            if isinstance(col, str) and col.lower() == 'vehicle class':
                vehicle_class_col = col
                break

        if vehicle_class_col is None:
            print(f"Error: Column 'Vehicle Class' not found in {filename}. Skipping this file.")
            continue

        # Selecting required columns and handling column variations
        column_mapping = {}
        for req_col in required_columns:
            for col in df.columns:
                if isinstance(col, str) and col.lower() == req_col.lower():
                    column_mapping[req_col] = col
                    break
            # Handling column name variations (e.g., March 2019 case)
            if req_col == 'First Registration Vehicle Status (Note)' and req_col not in column_mapping:
                for col in df.columns:
                    if isinstance(col, str) and col.lower() == 'first registration vehicle status':
                        column_mapping[req_col] = col
                        print(f"Mapping 'First Registration Vehicle Status' to 'First Registration Vehicle Status (Note)' for {filename}")
                        break

        missing_columns = [col for col in required_columns if col not in column_mapping]
        if missing_columns:
            print(f"Error: Columns {missing_columns} not found in {filename}. Skipping this file.")
            continue

        df = df[[column_mapping[col] for col in required_columns]]
        df.columns = required_columns

        # Converting '-' to NaN in 'Permitted Gross Vehicle Weight' column
        df['Permitted Gross Vehicle Weight'] = df['Permitted Gross Vehicle Weight'].replace('-', np.nan)

        # Filtering for Vehicle Class 'LGV' and Permitted Gross Vehicle Weight <= 5.5 (excluding NaN)
        filtered_df = df[
            (df['Vehicle Class'] == 'LGV') &
            (df['Permitted Gross Vehicle Weight'].notna()) &
            (df['Permitted Gross Vehicle Weight'].astype(float) <= 5.5)
        ].copy()  # Creating a copy to avoid SettingWithCopyWarning

        # Adding 'Month' and 'Registration Year' columns using .loc
        filtered_df.loc[:, 'Month'] = month
        filtered_df.loc[:, 'Registration Year'] = 2025

        # Appending the filtered DataFrame to the list
        all_dfs.append(filtered_df)
    except Exception as e:
        print(f"Error processing file {filename}: {e}")

# Combining all DataFrames into one
if all_dfs:
    combined_df = pd.concat(all_dfs, ignore_index=True)
    # Saving the combined data to a new Excel file
    combined_df.to_excel(output_file, index=False)
    print(f"\nProcessed data saved to {output_file}")
else:
    print("\nNo data was processed. Please check the input files and error messages above.")

# Logging contents of potentially problematic files (e.g., May) for review
for filename in expected_files:
    if 'may' in filename.lower():
        file_path = os.path.join(input_dir, filename)
        if os.path.exists(file_path):
            print(f"\nLogging data from {filename} for review due to potential format issues:")
            try:
                may_df = pd.read_excel(file_path, header=3)
                print("May 2025 file column names:", may_df.columns.tolist())
                print("May 2025 file first few rows:")
                print(may_df.head())
            except Exception as e:
                print(f"Error reading {filename}: {e}")