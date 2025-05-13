import pandas as pd
import os
import re

# 定義從檔案名稱中提取月份的函數 - 針對新格式優化
def extract_month_from_filename(filename):
    # 針對 "YYYYMM" 格式的檔案名稱
    match = re.match(r'^(\d{4})(\d{2}).*$', filename)
    if match:
        year = match.group(1)
        month = match.group(2)
        if year == '2018' and 1 <= int(month) <= 12:
            return month
    
    # 如果上面的方法失敗，嘗試其他方法
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
    
    # 嘗試使用正則表達式匹配月份
    for pattern, month_num in month_patterns:
        if re.search(pattern, filename_lower):
            return month_num
    
    # 嘗試從檔案名稱中提取數字
    digits = re.findall(r'\d+', filename)
    for digit in digits:
        if len(digit) <= 2 and 1 <= int(digit) <= 12:
            return f"{int(digit):02d}"
    
    return 'Unknown'

# 包含Excel檔案的目錄
input_dir = r'D:\Users\paulli1\Desktop\MINI BUS Research\Market Data\2018'
output_file = 'processed_light_bus_data_2018_all_months_complete.xlsx'

# 列出目錄中的所有檔案
print(f"目錄 {input_dir} 中的檔案：")
files_in_dir = os.listdir(input_dir)
for file in files_in_dir:
    print(f" - {file}")

# 初始化一個空列表來存儲DataFrame
all_dfs = []

# 建立月份與檔案的映射
month_to_files = {f"{i:02d}": [] for i in range(1, 13)}
unknown_files = []

# 先將檔案按月份分類
for file in files_in_dir:
    if not (file.endswith('.xlsx') or file.endswith('.xls')) or file.startswith('~$') or file == 'main.py':
        continue
    
    # 從檔案名稱中提取月份 - 針對新格式優化
    if file.startswith('2018'):
        # 直接從檔名提取月份
        month = file[4:6]
        if month.isdigit() and 1 <= int(month) <= 12:
            month_to_files[month].append(file)
            continue
    
    # 如果不是標準格式，嘗試其他方法
    month = extract_month_from_filename(file)
    
    if month != 'Unknown':
        month_to_files[month].append(file)
    else:
        unknown_files.append(file)

# 顯示每個月份的檔案
print("\n按月份分類的檔案：")
for month, files in month_to_files.items():
    print(f"月份 {month}: {len(files)} 個檔案")
    for file in files:
        print(f"  - {file}")

if unknown_files:
    print(f"\n無法識別月份的檔案: {len(unknown_files)} 個檔案")
    for file in unknown_files:
        print(f"  - {file}")
    
    # 手動分配月份
    print("\n請為無法識別月份的檔案手動分配月份：")
    for i, file in enumerate(unknown_files):
        print(f"{i+1}. {file}")
        print("   請輸入月份 (1-12)，或按Enter跳過：")
        try:
            user_input = input().strip()
            if user_input:
                month_num = int(user_input)
                if 1 <= month_num <= 12:
                    month = f"{month_num:02d}"
                    month_to_files[month].append(file)
                    print(f"   已將檔案 {file} 分配到月份 {month}")
                else:
                    print("   輸入的月份無效，檔案將被跳過")
            else:
                print("   已跳過此檔案")
        except:
            print("   輸入無效，檔案將被跳過")

# 定義所有需要的欄位
required_columns = [
    'Vehicle Class', 
    'Vehicle Make', 
    'Vehicle Model', 
    'Fuel Type',
    'Cylinder capacity of engine (c.c.)',
    'Body Type',
    'First Registration',
    'Vehicle Status (Note)',
    'Permitted Gross Vehicle Weight',
    'Number of passenger seats',
    'Taxable Value (HK$)',
    'Year Of Manufacture',
    'Month',
    'Registration Year'
]

# 處理每個月份的檔案
for month, files in month_to_files.items():
    if not files:  # 跳過沒有檔案的月份
        continue
        
    print(f"\n處理月份 {month} 的檔案：")
    month_num = int(month)
    
    for file in files:
        file_path = os.path.join(input_dir, file)
        print(f"\n處理檔案：{file}")
        
        try:
            # 根據檔案副檔名選擇適當的引擎
            engine = 'xlrd' if file.endswith('.xls') else None
            
            try:
                # 根據月份決定如何讀取檔案
                if month_num <= 5:  # 1-5月
                    print("檢測到1-5月的檔案格式，從第一列開始讀取資料...")
                    # 從第一列開始讀取，沒有標題列
                    df = pd.read_excel(file_path, header=None, engine=engine)
                    
                    # 列印前幾行以檢查資料
                    print("DataFrame的前幾行（無標題）：")
                    print(df.head())
                    
                    # 嘗試找到包含"Light Bus"的行
                    light_bus_rows = []
                    for i, row in df.iterrows():
                        row_str = ' '.join([str(val) for val in row.values if pd.notna(val)])
                        if 'light bus' in row_str.lower() or 'public light bus' in row_str.lower() or 'private light bus' in row_str.lower():
                            print(f"找到Light Bus資料，行號：{i}")
                            row_dict = {}
                            
                            # 嘗試映射欄位 - 根據檔案實際結構調整
                            col_mapping = {
                                0: 'Vehicle Class',
                                1: 'Vehicle Make',
                                2: 'Vehicle Model',
                                3: 'Fuel Type',
                                4: 'Cylinder capacity of engine (c.c.)',
                                5: 'Body Type',
                                6: 'First Registration',
                                7: 'Vehicle Status (Note)',
                                8: 'Permitted Gross Vehicle Weight',
                                9: 'Number of passenger seats',
                                10: 'Taxable Value (HK$)',
                                11: 'Year Of Manufacture'
                            }
                            
                            # 從行中提取數據
                            for col_idx, col_name in col_mapping.items():
                                if col_idx < len(row) and pd.notna(row[col_idx]):
                                    row_dict[col_name] = row[col_idx]
                                else:
                                    row_dict[col_name] = None
                            
                            # 添加月份和年份
                            row_dict['Month'] = month
                            row_dict['Registration Year'] = 2018
                            light_bus_rows.append(row_dict)
                    
                    if light_bus_rows:
                        # 創建一個新的DataFrame來存儲Light Bus資料
                        light_bus_df = pd.DataFrame(light_bus_rows)
                        
                        # 確保有所需的所有列
                        for col in required_columns:
                            if col not in light_bus_df.columns:
                                light_bus_df[col] = None
                        
                        # 只保留必要的列
                        light_bus_df = light_bus_df[required_columns]
                        
                        all_dfs.append(light_bus_df)
                        print(f"從 {file} 中提取了 {len(light_bus_df)} 筆Light Bus資料。")
                    else:
                        print(f"在 {file} 中沒有找到Light Bus資料。")
                else:  # 6-12月
                    print("檢測到6-12月的檔案格式，從第四列開始讀取資料...")
                    # 嘗試不同的header值來找到正確的標題行
                    for header_row in [3, 2, 1, 0]:
                        try:
                            df = pd.read_excel(file_path, header=header_row, engine=engine)
                            # 去除列名中的空白
                            df.columns = df.columns.str.strip()
                            print(f"嘗試header={header_row}，列名：", df.columns.tolist())
                            
                            # 檢查是否有'Vehicle Class'列
                            if 'Vehicle Class' in df.columns:
                                break
                        except Exception as e:
                            print(f"嘗試header={header_row}時出錯：{e}")
                    
                    # 列印前幾行以檢查資料
                    print("DataFrame的前幾行：")
                    print(df.head())
                    
                    # 檢查是否有'Vehicle Class'列
                    if 'Vehicle Class' in df.columns:
                        # 過濾Vehicle Class為'Private Light Bus'或'Public Light Bus'
                        filtered_df = df[df['Vehicle Class'].str.contains('Light Bus', case=False, na=False)]
                        if len(filtered_df) > 0:
                            # 添加'Month'和'Registration Year'列
                            filtered_df = filtered_df.copy()  # 避免SettingWithCopyWarning
                            
                            # 確保有所需的所有列
                            for col in required_columns:
                                if col not in filtered_df.columns and col not in ['Month', 'Registration Year']:
                                    filtered_df[col] = None
                            
                            filtered_df['Month'] = month
                            filtered_df['Registration Year'] = 2018
                            
                            # 只保留必要的列
                            available_columns = [col for col in required_columns if col in filtered_df.columns]
                            filtered_df = filtered_df[available_columns]
                            
                            all_dfs.append(filtered_df)
                            print(f"從 {file} 中提取了 {len(filtered_df)} 筆Light Bus資料。")
                        else:
                            print(f"在 {file} 中沒有找到Light Bus資料。")
                    else:
                        print(f"在 {file} 中找不到'Vehicle Class'列。")
                        # 嘗試在所有列中尋找包含"Light Bus"的資料
                        light_bus_rows = []
                        for i, row in df.iterrows():
                            row_str = ' '.join([str(val) for val in row.values if pd.notna(val)])
                            if 'light bus' in row_str.lower() or 'public light bus' in row_str.lower() or 'private light bus' in row_str.lower():
                                print(f"找到Light Bus資料，行號：{i}")
                                row_dict = {}
                                
                                # 嘗試從列名中映射欄位
                                for j, col_name in enumerate(df.columns):
                                    if j < len(row) and pd.notna(row[j]):
                                        # 嘗試將列名映射到所需欄位
                                        mapped_col = None
                                        col_lower = str(col_name).lower()
                                        
                                        if 'class' in col_lower:
                                            mapped_col = 'Vehicle Class'
                                        elif 'make' in col_lower:
                                            mapped_col = 'Vehicle Make'
                                        elif 'model' in col_lower:
                                            mapped_col = 'Vehicle Model'
                                        elif 'fuel' in col_lower or 'type' in col_lower:
                                            mapped_col = 'Fuel Type'
                                        elif 'cylinder' in col_lower or 'capacity' in col_lower or 'engine' in col_lower or 'c.c' in col_lower:
                                            mapped_col = 'Cylinder capacity of engine (c.c.)'
                                        elif 'body' in col_lower:
                                            mapped_col = 'Body Type'
                                        elif 'first' in col_lower and 'registration' in col_lower:
                                            mapped_col = 'First Registration'
                                        elif 'status' in col_lower:
                                            mapped_col = 'Vehicle Status (Note)'
                                        elif 'weight' in col_lower or 'gross' in col_lower:
                                            mapped_col = 'Permitted Gross Vehicle Weight'
                                        elif 'seat' in col_lower or 'passenger' in col_lower:
                                            mapped_col = 'Number of passenger seats'
                                        elif 'value' in col_lower or 'taxable' in col_lower or 'hk$' in col_lower:
                                            mapped_col = 'Taxable Value (HK$)'
                                        elif 'year' in col_lower and 'manufacture' in col_lower:
                                            mapped_col = 'Year Of Manufacture'
                                        else:
                                            # 如果找不到匹配，使用原始列名
                                            mapped_col = str(col_name)
                                        
                                        row_dict[mapped_col] = row[j]
                                
                                # 如果沒有找到Vehicle Class，添加一個
                                if 'Vehicle Class' not in row_dict:
                                    if 'public light bus' in row_str.lower():
                                        row_dict['Vehicle Class'] = 'Public Light Bus'
                                    elif 'private light bus' in row_str.lower():
                                        row_dict['Vehicle Class'] = 'Private Light Bus'
                                    else:
                                        row_dict['Vehicle Class'] = 'Light Bus'
                                
                                # 添加月份和年份
                                row_dict['Month'] = month
                                row_dict['Registration Year'] = 2018
                                light_bus_rows.append(row_dict)
                        
                        if light_bus_rows:
                            light_bus_df = pd.DataFrame(light_bus_rows)
                            
                            # 確保有所需的所有列
                            for col in required_columns:
                                if col not in light_bus_df.columns:
                                    light_bus_df[col] = None
                            
                            # 只保留必要的列
                            available_columns = [col for col in required_columns if col in light_bus_df.columns]
                            light_bus_df = light_bus_df[available_columns]
                            
                            all_dfs.append(light_bus_df)
                            print(f"從 {file} 中提取了 {len(light_bus_df)} 筆Light Bus資料。")
                        else:
                            print(f"在 {file} 中沒有找到Light Bus資料。")
            except Exception as e:
                print(f"讀取檔案時出錯：{e}")
                print("嘗試安裝xlrd庫...")
                import subprocess
                try:
                    subprocess.check_call(['pip', 'install', 'xlrd>=2.0.1'])
                    print("xlrd庫安裝成功，再次嘗試讀取檔案...")
                    # 重新嘗試讀取檔案
                    if month_num <= 5:
                        df = pd.read_excel(file_path, header=None, engine='xlrd')
                    else:
                        df = pd.read_excel(file_path, header=3, engine='xlrd')
                    print("成功讀取檔案！")
                except Exception as e2:
                    print(f"安裝xlrd或再次讀取檔案時出錯：{e2}")
                    continue
        except Exception as e:
            print(f"處理檔案 {file} 時出錯：{e}")

# 將所有DataFrame合併為一個
if all_dfs:
    # 檢查每個DataFrame的列是否兼容
    print("\n檢查每個DataFrame的列...")
    for i, df in enumerate(all_dfs):
        print(f"DataFrame {i+1} 的列：{df.columns.tolist()}")
    
    # 標準化所有DataFrame的列
    standardized_dfs = []
    
    for i, df in enumerate(all_dfs):
        # 確保所有必要的列都存在
        for col in required_columns:
            if col not in df.columns:
                df[col] = None  # 如果列不存在，添加一個空列
        
        # 只保留必要的列
        df = df[required_columns]
        standardized_dfs.append(df)
        print(f"已標準化DataFrame {i+1}")
    
    # 嘗試合併所有標準化後的DataFrame
    try:
        combined_df = pd.concat(standardized_dfs, ignore_index=True)
        # 將合併後的資料保存到新的Excel檔案
        combined_df.to_excel(output_file, index=False)
        print(f"\n處理後的資料已保存到 {output_file}")
        print(f"總共找到 {len(combined_df)} 筆Light Bus資料。")
    except Exception as e:
        print(f"合併DataFrame時出錯：{e}")
        print("嘗試單獨保存每個DataFrame...")
        for i, df in enumerate(standardized_dfs):
            month_output_file = f'processed_light_bus_data_2018_month_{i+1}.xlsx'
            df.to_excel(month_output_file, index=False)
            print(f"已將DataFrame {i+1} 保存到 {month_output_file}")
else:
    print("\n沒有處理任何資料。請檢查輸入檔案和上述錯誤訊息。")
