import pandas as pd

# 1. 設計替換數據
df_replacements = pd.DataFrame({
    'Tag': ['{{Project}}', '{{Date}}', '{{Author}}'],
    'Value': ['Super Laptop X', '2024-05-07', 'Luke & Piero']
})

# 2. 設計測試狀態
df_status = pd.DataFrame({
    '測項名稱': ['1.1 Display Test', '1.2 Keyboard Test', '1.3 Audio Test'],
    '測試狀態': ['Pass', 'Skipped', 'Fail']
})

# 3. 儲存至同一個 Excel 的不同分頁
with pd.ExcelWriter('Report_Config.xlsx') as writer:
    df_replacements.to_excel(writer, sheet_name='Replacements', index=False)
    df_status.to_excel(writer, sheet_name='Test_Status', index=False)

print('Report_Config.xlsx 已成功創建於當前目錄。')
