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
    # Extracting month name from text like "Last updated: September 2019"
    if isinstance(text, str):
        match = re.search(r'(january|february|march|april|may|june|july|august|september|october|november|december)', text.lower(), re.IGNORECASE)
        if match:
            month_name = match.group(1).lower()
            return month_map.get(month_name, 'Unknown')
    print(f"Warning: Could not extract month from text: {text}")
    return 'Unknown'

# Directory containing the Excel files
input_dir = r'd:\Users\paulli1\Desktop\Project Flexis\Market Data\2019'
output_file = 'processed_light_bus_data_2019_all_months.xlsx'

# List of expected filenames for all months
expected_files = [
    'particulars of first registered vehicle_january 2019.xlsx',
    'particulars of first registered vehicle_february 2019.xlsx',
    'particulars of first registered vehicle_march 2019.xlsx',
    'particulars of first registered vehicle_april 2019.xlsx',
    'particulars of first registered vehicle_may 2019.xlsx',
    'particulars of first registered vehicle_june 2019.xlsx',
    'particulars of first registered vehicle_july 2019.xlsx',
    'particulars of first registered vehicle_august 2019.xlsx',
    'particulars of first registered vehicle_september 2019.xlsx',
    'particulars of first registered vehicle_october 2019.xlsx',
    'particulars of first registered vehicle_november 2019.xlsx',
    'particulars of first registered vehicle_december 2019.xlsx'
]

# Initializing an empty list to store DataFrames
all_dfs = []

# List of required columns
required_columns = [
    'Vehicle Class', 'Vehicle Make', 'Vehicle Model', 'Fuel Type', 
    'Cylinder capacity of engine (c.c.)', 'Body Type', 
    'First Registration Vehicle Status (Note)', 'Permitted Gross Vehicle Weight',
    'Number of passenger seats', 'Taxable Value (HK$)', 'Year Of Manufacture'
]

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

    # Selecting only the required columns with flexible mapping
    column_mapping = {}
    for req_col in required_columns:
        for col in df.columns:
            if col.lower() == req_col.lower():
                column_mapping[req_col] = col
                break
        # Handle March 2019 case: map 'First Registration Vehicle Status' to 'First Registration Vehicle Status (Note)'
        if req_col == 'First Registration Vehicle Status (Note)' and req_col not in column_mapping:
            for col in df.columns:
                if col.lower() == 'first registration vehicle status':
                    column_mapping[req_col] = col
                    print(f"Mapping 'First Registration Vehicle Status' to 'First Registration Vehicle Status (Note)' for {filename}")
                    break

    missing_columns = [col for col in required_columns if col not in column_mapping]
    if missing_columns:
        print(f"Error: Columns {missing_columns} not found in {filename}. Skipping this file.")
        continue

    df = df[[column_mapping[col] for col in required_columns]]
    df.columns = required_columns

    # Convert '-' to NaN in 'Permitted Gross Vehicle Weight' column
    df['Permitted Gross Vehicle Weight'] = df['Permitted Gross Vehicle Weight'].replace('-', np.nan)

    # Filtering for Vehicle Class 'LGV' and Permitted Gross Vehicle Weight <= 5.5 (excluding NaN)
    filtered_df = df[
        (df['Vehicle Class'] == 'LGV') & 
        (df['Permitted Gross Vehicle Weight'].notna()) & 
        (df['Permitted Gross Vehicle Weight'].astype(float) <= 5.5)
    ].copy()  # Create a copy to avoid SettingWithCopyWarning

    # Adding 'Month' and 'Registration Year' columns using .loc
    filtered_df.loc[:, 'Month'] = month
    filtered_df.loc[:, 'Registration Year'] = 2019

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

# Log May 2019 data separately for review
may_file = 'particulars of first registered vehicle_may 2019.xlsx'
may_file_path = os.path.join(input_dir, may_file)
if os.path.exists(may_file_path):
    print(f"\nLogging data from {may_file} for review due to incorrect format:")
    may_df = pd.read_excel(may_file_path, header=3)
    print("Column names in May 2019 file:", may_df.columns.tolist())
    print("First few rows of May 2019 file:")
    print(may_df.head())