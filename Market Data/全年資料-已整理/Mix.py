import pandas as pd
import os
import re

# Directory containing the Excel files
input_dir = r'D:\Users\paulli1\Desktop\Project Flexis\Market Data\全年資料-已整理'

# File name list (assuming processed_LGV_data_YYYY_all_months.xlsx based on previous naming)
files = [
    "processed_LGV_data_2019_all_months",
    "processed_LGV_data_2020_all_months",
    "processed_LGV_data_2021_all_months",
    "processed_LGV_data_2022_all_months",
    "processed_LGV_data_2023_all_months",
    "processed_LGV_data_2024_all_months",
    "processed_LGV_data_2025_all_months"
]

# Required columns list with exact names as specified
required_columns = [
    "Vehicle Class",
    "Vehicle Make",
    "Vehicle Model",
    "Fuel Type",
    "Cylinder Capacity Of Engine (c.c.)",
    "Rated Power (kW)",
    "Body Type",
    "First Registration Vehicle Status (Note)",
    "Permitted Gross Vehicle Weight",
    "Number Of Passenger Seats",
    "Taxable Value (HK$)",
    "Year Of Manufacture",
    "Month",
    "Registration Year"
]

# Create an empty DataFrame to store all combined data
all_data = pd.DataFrame()

# Process each file
for file in files:
    try:
        # Construct full file path
        file_path = os.path.join(input_dir, f"{file}.xlsx")
        
        # Extract year from filename
        year_match = re.search(r'(\d{4})', file)
        year = year_match.group(1) if year_match else "Unknown Year"
        
        print(f"Processing file: {file}.xlsx")
        
        # Read Excel file
        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
        else:
            print(f"Warning: File {file_path} not found. Skipping.")
            continue
        
        # Check and retain only required columns
        available_columns = [col for col in required_columns if col in df.columns]
        if len(available_columns) < len(required_columns):
            missing_columns = [col for col in required_columns if col not in df.columns]
            print(f"Warning: File {file}.xlsx is missing the following columns: {missing_columns}")
            
            # Display actual columns in the file for comparison
            print(f"Actual columns in the file: {list(df.columns)}")
        
        df_filtered = df[available_columns].copy()
        
        # Add 'Registration Year' if it doesn't exist, using the extracted year
        if "Registration Year" not in df_filtered.columns:
            df_filtered["Registration Year"] = year
        
        # Add missing required columns with None values
        for col in required_columns:
            if col not in df_filtered.columns:
                df_filtered[col] = None
        
        # Ensure all required columns are present in the order specified
        df_filtered = df_filtered[required_columns]
        
        # Concatenate the processed data to the combined DataFrame
        all_data = pd.concat([all_data, df_filtered], ignore_index=True)
        
    except Exception as e:
        print(f"Error processing file {file}.xlsx: {str(e)}")

# Check if any data was combined
if all_data.empty:
    print("No data was combined. Please check file paths and file contents.")
else:
    # Save as Excel file
    excel_output = "processed_LGV_data_2019-2025.xlsx"
    all_data.to_excel(excel_output, index=False)
    print(f"Combined data saved as Excel file: {excel_output}")
    
    # Display combined data basic information
    print(f"\nCombined data contains {len(all_data)} rows and {len(all_data.columns)} columns")
    print("Column list:")
    for col in all_data.columns:
        print(f"- {col}")