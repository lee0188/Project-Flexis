import pandas as pd
import os

# 檔案名稱列表
files = [
    "processed_light_bus_data_2018_all_months",
    "processed_light_bus_data_2019_all_months",
    "processed_light_bus_data_2020_all_months",
    "processed_light_bus_data_2021_all_months",
    "processed_light_bus_data_2022_all_months",
    "processed_light_bus_data_2023_all_months",
    "processed_light_bus_data_2024_all_months",
    "processed_light_bus_data_2025_all_months"
]

# 需要保留的欄位列表，使用您提供的確切欄位名稱
required_columns = [
    "Vehicle Class",
    "Vehicle Make",
    "Vehicle Model",
    "Fuel Type",
    "Cylinder capacity of engine (c.c.)",  # 注意這裡的大小寫和格式
    "Body Type",
    "First Registration",
    "Vehicle Status (Note)",  # 這是一個單獨的欄位，不是 "First Registration Vehicle Status (Note)"
    "Permitted Gross Vehicle Weight",
    "Number of passenger seats",  # 注意這裡的大小寫
    "Taxable Value (HK$)",
    "Year Of Manufacture",
    "Month",
    "Registration Year"
]

# 創建一個空的 DataFrame 來存儲所有合併的數據
all_data = pd.DataFrame()

# 處理每個檔案
for file in files:
    try:
        # 從檔案名稱中提取年份
        import re
        year_match = re.search(r'(\d{4})', file)
        year = year_match.group(1) if year_match else "未知年份"
        
        print(f"處理檔案: {file}.xlsx")
        
        # 讀取 Excel 檔案
        df = pd.read_excel(f"{file}.xlsx")
        
        # 檢查並只保留需要的欄位
        available_columns = [col for col in required_columns if col in df.columns]
        if len(available_columns) < len(required_columns):
            missing_columns = [col for col in required_columns if col not in df.columns]
            print(f"警告: 檔案 {file}.xlsx 缺少以下欄位: {missing_columns}")
            
            # 顯示檔案中實際的欄位，以便於比較
            print(f"檔案中的實際欄位: {list(df.columns)}")
        
        df_filtered = df[available_columns].copy()
        
        # 如果 "Registration Year" 欄位不存在，則添加一個並設置為檔案名稱中的年份
        if "Registration Year" not in df_filtered.columns:
            df_filtered["Registration Year"] = year
        
        # 將處理後的數據添加到合併的 DataFrame 中
        all_data = pd.concat([all_data, df_filtered], ignore_index=True)
        
    except Exception as e:
        print(f"處理檔案 {file}.xlsx 時發生錯誤: {str(e)}")

# 檢查是否有數據被合併
if all_data.empty:
    print("沒有數據被合併，請檢查文件路徑和文件內容。")
else:
    # 保存為 Excel 檔案
    excel_output = "merged_light_bus_data.xlsx"
    all_data.to_excel(excel_output, index=False)
    print(f"合併的數據已保存為 Excel 檔案: {excel_output}")
    
    # 保存為 CSV 檔案
    csv_output = "merged_light_bus_data.csv"
    all_data.to_csv(csv_output, index=False)
    print(f"合併的數據已保存為 CSV 檔案: {csv_output}")
    
    # 顯示合併後的數據基本信息
    print(f"\n合併後的數據包含 {len(all_data)} 行和 {len(all_data.columns)} 列")
    print("欄位列表:")
    for col in all_data.columns:
        print(f"- {col}")
