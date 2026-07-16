# -*- coding: utf-8 -*-
"""
Created on Fri Jul 10 20:02:44 2026

@author: Five-seveN
"""

import pandas as pd

# 1. 讀取原始檔案 (請確保檔案路徑正確)
# 這裡以您提供的 CSV 檔名為準，若未來直接讀取 Excel 檔，可改為 pd.read_excel('對照表.xlsx', sheet_name='Sheet3')
df_ref = pd.read_excel('對照表.xlsx', sheet_name='Sheet3')
df_data = pd.read_excel('灌輸.xlsx', sheet_name='Sheet1')

# 2. 清理「公司編號」欄位（去除文字前後的空白字元，以利精準比對）
df_ref['公司編號_clean'] = df_ref['公司編號'].astype(str).str.strip()
df_data['公司編號_clean'] = df_data['公司編號'].astype(str).str.strip()

# 3. 過濾掉 灌輸.xlsx 中重複出現的標題列 (公司編號欄位的值等於 '公司編號' 的列)
df_data_clean = df_data[df_data['公司編號_clean'] != '公司編號'].copy()

# 4. 去除重複的公司編號（若有重複代碼則保留第一筆），確保 1 對 1 對照，不影響原表列數
df_data_unique = df_data_clean.drop_duplicates(subset=['公司編號_clean'], keep='first')

# 5. 建立對照字典 (Key: 公司編號 -> Value: 記帳費 / 備註)
fee_map = df_data_unique.set_index('公司編號_clean')['記帳費']
note_map = df_data_unique.set_index('公司編號_clean')['備註']

# 6. 將對應的資料填入對照表的第 4 與第 5 直排，並設定欄位名稱
df_ref['記帳費'] = df_ref['公司編號_clean'].map(fee_map)
df_ref['備註'] = df_ref['公司編號_clean'].map(note_map)

# 7. 移除用於比對的暫存欄位
df_ref_final = df_ref.drop(columns=['公司編號_clean'])

# 8. 儲存結果 (使用 utf-8-sig 編碼以確保 Excel 開啟時中文不會變亂碼)
df_ref_final.to_csv('對照表_更新.csv', index=False, encoding='utf-8-sig')

print("對照與填入完成！已成功儲存為 '對照表_更新.csv'")