import pandas as pd
import os

# 1. 設計替換數據
df_replacements = pd.DataFrame({
    'Tag': ['{{Project}}', '{{Date}}', '{{Author}}'],
    'Value': ['Super Laptop X', '2024-05-07', 'Luke & Piero']
})

# 2. 設計測試狀態
test_items = [
    '1. Revision History',
    '2. Introduction',
    '2.1. Purpose',
    '2.2. Scope',
    '2.3. UUT Configuration',
    '2.4. Test Equipment and Software',
    '3. Summary',
    '3.1. Test Item Result',
    '3.2. Test Conclusion',
    '3.3. Bug List',
    '3.4. UUT Photo',
    '4. Reliability Test Result',
    '4.1. Environmental Qualification Test',
    '4.1.1. Non-Operational Storage Test',
    '4.1.2. Operational Temperature Cycle Test',
    '4.1.3. Cold / Hot Start Test',
    '4.2. Mechanical Structure Test',
    '4.2.1. Operating Vibration Test',
    '4.2.2. Non-operating Vibration Test',
    '4.2.3. Operating Shock Test',
    '4.2.4. Non-Operation Square WaveShock Test Report',
    '4.3. Packaged Test',
    '4.3.1. Package Test',
    '4.4. Other Test',
    '4.4.1. Rotational Vibration Index Test'
]

df_status = pd.DataFrame({
    '測項名稱': test_items,
    '測試狀態': ['Pass'] * len(test_items)
})

# 3. 儲存至同一個 Excel 的不同分頁
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
output_path = os.path.join(BASE_DIR, 'Report_Config.xlsx')
with pd.ExcelWriter(output_path) as writer:
    df_replacements.to_excel(writer, sheet_name='Replacements', index=False)
    df_status.to_excel(writer, sheet_name='Test_Status', index=False)

print(f'Report_Config.xlsx 已成功創建於: {os.path.abspath(output_path)}')
