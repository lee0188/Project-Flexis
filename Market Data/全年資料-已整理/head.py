import pandas as pd
import os

# 檔案名稱列表 (從圖片中獲取)
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

# 創建一個空列表來存儲結果
results = []

# 處理每個檔案
for file in files:
    # 正確從檔案名稱中提取年份
    # 使用字符串分割和正則表達式來提取年份
    import re
    year_match = re.search(r'(\d{4})', file)
    if year_match:
        year = year_match.group(1)  # 提取匹配的年份
    else:
        year = "未知年份"
    
    try:
        # 讀取 Excel 檔案，假設檔案在當前目錄中
        # 注意：您可能需要調整檔案路徑
        df = pd.read_excel(f"{file}.xlsx")
        
        # 獲取第一列的資料（列標題）
        headers = list(df.columns)
        first_row_data = ", ".join(headers)
        
        # 將年份和資料添加到結果列表中
        results.append([year, first_row_data])
        
    except Exception as e:
        # 如果發生錯誤，記錄錯誤訊息
        results.append([year, f"錯誤: {str(e)}"])

# 創建一個 DataFrame 來存儲結果
result_df = pd.DataFrame(results, columns=["Year", "資料"])

# 將結果保存為 Excel 檔案
output_file = "first_row_data_summary.xlsx"
result_df.to_excel(output_file, index=False)

print(f"結果已保存至 {output_file}")
