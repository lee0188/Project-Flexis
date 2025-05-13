import pandas as pd
import os
import re

# 定義從"Last updated"文字中提取月份的函數
def extract_month_from_text(text):
    month_map = {
        'january': '01', 'jan': '01', 'february': '02', 'feb': '02', 'march': '03', 'mar': '03',
        'april': '04', 'apr': '04', 'may': '05', 'june': '06', 'jun': '06', 'july': '07', 'jul': '07',
        'august': '08', 'aug': '08', 'september': '09', 'sep': '09', 'sept': '09', 
        'october': '10', 'oct': '10', 'november': '11', 'nov': '11', 'december': '12', 'dec': '12'
    }
    # 從文字中提取月份名稱，例如 "Last updated: September 2019"
    if isinstance(text, str):
        match = re.search(r'(january|february|march|april|may|june|july|august|september|october|november|december)', text.lower(), re.IGNORECASE)
        if match:
            month_name = match.group(1).lower()
            return month_map.get(month_name, 'Unknown')
    print(f"警告：無法從文字中提取月份：{text}")
    return 'Unknown'

# 從檔案名稱中提取月份
def extract_month_from_filename(filename):
    month_patterns = {
        r'january': '01', r'february': '02', r'march': '03', r'april': '04',
        r'may': '05', r'june': '06', r'july': '07', r'aug': '08',
        r'sept': '09', r'oct': '10', r'nov': '11', r'dec': '12'
    }
    
    filename_lower = filename.lower()
    for pattern, month_num in month_patterns.items():
        if pattern in filename_lower:
            return month_num
    
    print(f"警告：無法從檔案名稱中提取月份：{filename}")
    return 'Unknown'

# 包含Excel檔案的目錄
input_dir = r'D:\Users\paulli1\Desktop\MINI BUS Research\Market Data\2020'
output_file = 'processed_light_bus_data_2020_all_months.xlsx'

# 2020年的檔案名稱，根據圖片中的檔案名稱
expected_files = [
    'particulars of first registered vehicle_january 2020.xlsx',
    'particulars of first registered vehicle_february 2020.xlsx',
    'particulars of first registered vehicle_march 2020.xlsx',
    'particulars of first registered vehicle_april 2020.xlsx',
    'particulars of first registered vehicle_may 2020.xlsx',
    'particulars of first registered vehicle_june2020.xlsx',
    'particulars of first registered vehicle_july2020.xlsx',
    'particulars of first registered vehicle_aug2020.xlsx',
    'particulars of first registered vehicle_sept2020.xlsx',
    'particulars of first registered vehicle_oct2020.xlsx',
    'particulars of first registered vehicle_nov2020.xlsx',
    'particulars of first registered vehicle_dec2020.xlsx'
]

# 列出目錄中的所有檔案以進行調試
print(f"目錄 {input_dir} 中的檔案：")
files_in_dir = os.listdir(input_dir)
for file in files_in_dir:
    print(f" - {file}")

# 初始化一個空列表來存儲DataFrame
all_dfs = []

# 處理每個預期檔案
for filename in expected_files:
    file_path = os.path.join(input_dir, filename)
    if not os.path.exists(file_path):
        print(f"警告：找不到檔案：{filename}。跳過。")
        continue

    print(f"\n處理檔案：{filename}")
    
    # 從檔案名稱中提取月份
    month = extract_month_from_filename(filename)
    print(f"從檔案名稱提取的月份：{month}")

    try:
        # 嘗試讀取第二行以獲取"Last updated"文字
        try:
            second_row = pd.read_excel(file_path, header=None, skiprows=1, nrows=1)
            last_updated_text = second_row.iloc[0, 0]  # 假設文字在第一列
            month_from_text = extract_month_from_text(last_updated_text)
            if month_from_text != 'Unknown':
                month = month_from_text
                print(f"從'Last updated'文字提取的月份：{month}")
        except Exception as e:
            print(f"讀取'Last updated'文字時出錯：{e}，將使用從檔案名稱提取的月份")
        
        # 從第四行開始讀取主要資料（基於0的索引為3）
        df = pd.read_excel(file_path, header=3)

        # 去除列名中的空白
        df.columns = df.columns.str.strip()

        # 列印列名以進行調試
        print("Excel檔案中的列名：", df.columns.tolist())

        # 列印前幾行以檢查資料
        print("DataFrame的前幾行：")
        print(df.head())

        # 找到'Vehicle Class'列（不區分大小寫）
        vehicle_class_col = None
        for col in df.columns:
            if isinstance(col, str) and col.lower() == 'vehicle class':
                vehicle_class_col = col
                break

        if vehicle_class_col is None:
            print(f"錯誤：在 {filename} 中找不到'Vehicle Class'列。跳過此檔案。")
            continue

        # 過濾Vehicle Class為'Private Light Bus'或'Public Light Bus'
        filtered_df = df[df[vehicle_class_col].isin(['Private Light Bus', 'Public Light Bus'])]

        # 添加'Month'和'Registration Year'列
        filtered_df['Month'] = month
        filtered_df['Registration Year'] = 2020

        # 將過濾後的DataFrame添加到列表
        all_dfs.append(filtered_df)
    except Exception as e:
        print(f"處理檔案 {filename} 時出錯：{e}")

# 將所有DataFrame合併為一個
if all_dfs:
    combined_df = pd.concat(all_dfs, ignore_index=True)
    # 將合併後的資料保存到新的Excel檔案
    combined_df.to_excel(output_file, index=False)
    print(f"\n處理後的資料已保存到 {output_file}")
else:
    print("\n沒有處理任何資料。請檢查輸入檔案和上述錯誤訊息。")
