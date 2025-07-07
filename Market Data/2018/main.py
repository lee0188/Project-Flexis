import pandas as pd
import os
import re
import numpy as np

# Set pandas option to avoid FutureWarning for downcasting
pd.set_option('future.no_silent_downcasting', True)

# Defining function to extract month from filename - optimized for new format
def extract_month_from_filename(filename):
    # For "YYYYMM" format filenames
    match = re.match(r'^(\d{4})(\d{2}).*$', filename)
    if match:
        year = match.group(1)
        month = match.group(2)
        if year == '2018' and 1 <= int(month) <= 12:
            return month
    
    # Fallback method using month patterns
    month_patterns = [
        (r'jan(?:uary)?', '01'),
        (r'feb(?:ruary)?', '02'),
        (r'mar(?:ch)?', '03'),
        (r'apr(?:il)?', '04'),
        (r'may', '05'),
        (r'jun(?:e)?', '06'),
        (r'jul(?:y)?', '07'),
        (r'aug(?:ust)?', '08'),
        (r'sep(?:tember)?', '09'),
        (r'oct(?:ober)?', '10'),
        (r'nov(?:ember)?', '11'),
        (r'dec(?:ember)?', '12')
    ]
    
    filename_lower = filename.lower()
    for pattern, month_num in month_patterns:
        if re.search(pattern, filename_lower):
            return month_num
    
    # Fallback to extract digits
    digits = re.findall(r'\d+', filename)
    for digit in digits:
        if len(digit) <= 2 and 1 <= int(digit) <= 12:
            return f"{int(digit):02d}"
    
    return 'Unknown'

# Directory containing the Excel files
input_dir = r'D:\Users\paulli1\Desktop\Project Flexis\Market Data\2018'
output_file = 'processed_LGV_data_2018_all_months.xlsx'

# Listing all files in the directory for debugging
print(f"Files in directory {input_dir}:")
files_in_dir = os.listdir(input_dir)
for file in files_in_dir:
    print(f" - {file}")

# Initializing an empty list to store DataFrames
all_dfs = []

# Creating a mapping of months to files
month_to_files = {f"{i:02d}": [] for i in range(1, 13)}
unknown_files = []

# Classifying files by month
for file in files_in_dir:
    if not (file.endswith('.xlsx') or file.endswith('.xls')) or file.startswith('~$') or file == 'main.py':
        continue
    
    month = extract_month_from_filename(file)
    if month != 'Unknown':
        month_to_files[month].append(file)
    else:
        unknown_files.append(file)

# Displaying files classified by month
print("\nFiles classified by month:")
for month, files in month_to_files.items():
    print(f"Month {month}: {len(files)} files")
    for file in files:
        print(f"  - {file}")

if unknown_files:
    print(f"\nFiles with unidentifiable months: {len(unknown_files)} files")
    for file in unknown_files:
        print(f"  - {file}")

# Required columns
required_columns = [
    'Vehicle Class', 'Vehicle Make', 'Vehicle Model', 'Fuel Type',
    'Cylinder capacity of engine (c.c.)', 'Body Type',
    'First Registration Vehicle Status (Note)', 'Permitted Gross Vehicle Weight',
    'Number of passenger seats', 'Taxable Value (HK$)', 'Year Of Manufacture'
]

# Processing each month's files
for month, files in month_to_files.items():
    if not files:  # Skip months with no files
        continue
        
    print(f"\nProcessing files for month {month}:")
    month_num = int(month)
    
    for file in files:
        file_path = os.path.join(input_dir, file)
        print(f"\nProcessing file: {file}")
        
        try:
            # Select appropriate engine based on file extension
            engine = 'xlrd' if file.endswith('.xls') else None
            
            try:
                # Determine how to read the file based on month
                if month_num <= 5:  # January–May
                    print("Detected January–May file format, reading data from the first row...")
                    df = pd.read_excel(file_path, header=None, engine=engine)
                    
                    # Print first few rows to inspect data
                    print("First few rows of the DataFrame (no headers):")
                    print(df.head())
                    
                    # Attempt to find rows containing "Light Goods Vehicle"
                    light_vehicle_rows = []
                    for i, row in df.iterrows():
                        row_str = ' '.join([str(val) for val in row.values if pd.notna(val)])
                        if 'light goods vehicle' in row_str.lower():
                            print(f"Found Light Goods Vehicle data, row number: {i}")
                            row_dict = {}
                            
                            # Map columns based on file structure, with enhanced 'Year Of Manufacture' detection
                            col_mapping = {
                                0: 'Vehicle Class',
                                1: 'Vehicle Make',
                                2: 'Vehicle Model',
                                3: 'Fuel Type',
                                4: 'Cylinder capacity of engine (c.c.)',
                                5: 'Body Type',
                                6: 'First Registration Vehicle Status (Note)',
                                8: 'Permitted Gross Vehicle Weight',
                                9: 'Number of passenger seats',
                                10: 'Taxable Value (HK$)',
                                11: 'Year Of Manufacture'  # Assuming this is the typical position
                            }
                            
                            # Extract data from the row
                            for col_idx, col_name in col_mapping.items():
                                if col_idx < len(row) and pd.notna(row[col_idx]):
                                    # Special handling for 'Year Of Manufacture' to detect numeric years
                                    if col_name == 'Year Of Manufacture' and isinstance(row[col_idx], (int, float)):
                                        row_dict[col_name] = int(row[col_idx]) if row[col_idx].is_integer() else None
                                    elif col_name == 'Year Of Manufacture' and isinstance(row[col_idx], str):
                                        year_match = re.search(r'\b(19|20)\d{2}\b', str(row[col_idx]))
                                        row_dict[col_name] = int(year_match.group(0)) if year_match else None
                                    else:
                                        row_dict[col_name] = row[col_idx]
                                else:
                                    row_dict[col_name] = None
                            
                            # Add month and year
                            row_dict['Month'] = month
                            row_dict['Registration Year'] = 2018
                            light_vehicle_rows.append(row_dict)
                    
                    if light_vehicle_rows:
                        light_vehicle_df = pd.DataFrame(light_vehicle_rows)
                        
                        # Ensure all required columns are present
                        for col in required_columns:
                            if col not in light_vehicle_df.columns:
                                light_vehicle_df[col] = None
                        
                        # Keep only required columns
                        light_vehicle_df = light_vehicle_df[required_columns]
                        all_dfs.append(light_vehicle_df)
                        print(f"Extracted {len(light_vehicle_df)} Light Goods Vehicle records from {file}.")
                    else:
                        print(f"No Light Goods Vehicle data found in {file}.")
                else:  # June–December
                    print("Detected June–December file format, reading data from the fourth row...")
                    # Try different header values to find the correct header row
                    for header_row in [3, 2, 1, 0]:
                        try:
                            df = pd.read_excel(file_path, header=header_row, engine=engine)
                            # Strip whitespace from column names
                            df.columns = df.columns.str.strip()
                            print(f"Tried header={header_row}, column names:", df.columns.tolist())
                            
                            # Check if 'Vehicle Class' column exists
                            if 'Vehicle Class' in df.columns:
                                break
                        except Exception as e:
                            print(f"Error with header={header_row}: {e}")
                    
                    # Print first few rows to inspect data
                    print("First few rows of the DataFrame:")
                    print(df.head())
                    
                    # Check if 'Vehicle Class' column exists
                    if 'Vehicle Class' in df.columns:
                        # Convert '-' to NaN in 'Permitted Gross Vehicle Weight'
                        df['Permitted Gross Vehicle Weight'] = df['Permitted Gross Vehicle Weight'].replace('-', np.nan)
                        
                        # Filter for Vehicle Class 'Light Goods Vehicle' and Permitted Gross Vehicle Weight <= 5.5
                        filtered_df = df[
                            (df['Vehicle Class'] == 'Light Goods Vehicle') &
                            (df['Permitted Gross Vehicle Weight'].notna()) &
                            (df['Permitted Gross Vehicle Weight'].astype(float) <= 5.5)
                        ].copy()  # Create a copy to avoid SettingWithCopyWarning
                        
                        if len(filtered_df) > 0:
                            # Enhance 'Year Of Manufacture' extraction
                            if 'Year Of Manufacture' not in filtered_df.columns:
                                # Attempt to detect 'Year Of Manufacture' based on numeric values
                                for col in df.columns:
                                    if col.lower().startswith('year') or col.lower().startswith('manufacture'):
                                        filtered_df['Year Of Manufacture'] = df[col].where(
                                            df[col].apply(lambda x: isinstance(x, (int, float)) and 1900 <= x <= 2018),
                                            None
                                        )
                                        break
                                    elif df[col].dtype in [int, float]:
                                        filtered_df['Year Of Manufacture'] = df[col].where(
                                            df[col].apply(lambda x: isinstance(x, (int, float)) and 1900 <= x <= 2018),
                                            None
                                        )
                                        break
                            
                            # Ensure all required columns are present
                            for col in required_columns:
                                if col not in filtered_df.columns:
                                    filtered_df[col] = None
                            
                            # Add 'Month' and 'Registration Year' columns using .loc
                            filtered_df.loc[:, 'Month'] = month
                            filtered_df.loc[:, 'Registration Year'] = 2018
                            
                            # Keep only required columns
                            available_columns = [col for col in required_columns if col in filtered_df.columns]
                            filtered_df = filtered_df[available_columns]
                            
                            all_dfs.append(filtered_df)
                            print(f"Extracted {len(filtered_df)} Light Goods Vehicle records from {file}.")
                        else:
                            print(f"No Light Goods Vehicle data found in {file} with the specified weight condition.")
                    else:
                        print(f"Column 'Vehicle Class' not found in {file}.")
            except Exception as e:
                print(f"Error reading file: {e}")
                print("Attempting to install xlrd...")
                import subprocess
                try:
                    subprocess.check_call(['pip', 'install', 'xlrd>=2.0.1'])
                    print("xlrd installed successfully, retrying file read...")
                    if month_num <= 5:
                        df = pd.read_excel(file_path, header=None, engine='xlrd')
                    else:
                        df = pd.read_excel(file_path, header=3, engine='xlrd')
                    print("Successfully read file!")
                except Exception as e2:
                    print(f"Error installing xlrd or retrying file read: {e2}")
                    continue
        except Exception as e:
            print(f"Error processing file {file}: {e}")

# Combining all DataFrames into one
if all_dfs:
    # Check compatibility of columns in each DataFrame
    print("\nChecking columns in each DataFrame...")
    for i, df in enumerate(all_dfs):
        print(f"DataFrame {i+1} columns: {df.columns.tolist()}")
    
    # Standardize all DataFrames' columns
    standardized_dfs = []
    for i, df in enumerate(all_dfs):
        # Ensure all required columns are present
        for col in required_columns:
            if col not in df.columns:
                df[col] = None  # Add empty column if missing
        
        # Keep only required columns
        df = df[required_columns]
        standardized_dfs.append(df)
        print(f"Standardized DataFrame {i+1}")
    
    # Attempt to concatenate all standardized DataFrames
    try:
        combined_df = pd.concat(standardized_dfs, ignore_index=True)
        # Save the combined data to a new Excel file
        combined_df.to_excel(output_file, index=False)
        print(f"\nProcessed data saved to {output_file}")
        print(f"Total of {len(combined_df)} Light Goods Vehicle records found.")
    except Exception as e:
        print(f"Error concatenating DataFrames: {e}")
        print("Attempting to save each DataFrame separately...")
        for i, df in enumerate(standardized_dfs):
            month_output_file = f'processed_LGV_data_2018_month_{i+1}.xlsx'
            df.to_excel(month_output_file, index=False)
            print(f"Saved DataFrame {i+1} to {month_output_file}")
else:
    print("\nNo data was processed. Please check the input files and error messages above.")

# Logging contents of potentially problematic files (e.g., May) for review
for file in files_in_dir:
    if 'may' in file.lower() and (file.endswith('.xlsx') or file.endswith('.xls')):
        file_path = os.path.join(input_dir, file)
        if os.path.exists(file_path):
            print(f"\nLogging data from {file} for review due to potential format issues:")
            try:
                may_df = pd.read_excel(file_path, header=3)
                print("May 2018 file column names:", may_df.columns.tolist())
                print("May 2018 file first few rows:")
                print(may_df.head())
            except Exception as e:
                print(f"Error reading {file}: {e}")