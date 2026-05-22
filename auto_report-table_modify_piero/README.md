# 表格修改工具 (Auto Report Table Modify)

**版本**: v1.0  
**開發者**: Piero  
**狀態**: ✅ 生產版本

## 🎯 功能

用於修改 Word RT 報告中表格內容的工具。支援根據 Excel 配置自動更新表格單元格的數值、狀態等信息。

## ✨ 主要特性

- **表格值替換**: 自動根據配置替換表格中的數值
- **狀態更新**: 更新測試結果、狀態欄位
- **批量修改**: 支援批量處理多個表格
- **配置驅動**: 基於 Excel 配置文件進行操作
- **格式保留**: 修改時保留原始格式

## 📋 使用方式

### 環境需求
- Python 3.7+
- pandas
- pywin32
- openpyxl
- win32com

### 快速開始

```python
# 1. 準備 Excel 配置文件 (Master_Log.xlsx)
#    定義要修改的表格和對應的值

# 2. 準備 Word 範本 (Template.doc)

# 3. 執行腳本
python auto_report_table_modify.py
```

### 配置說明

**Excel 檔案格式**:
```
| 表格名稱 | 單元格位置 | 新值 | 備註 |
|---------|----------|------|------|
| Table 1 | B2 | PASS | 測試結果 |
| Table 1 | C2 | 2024-01-15 | 測試日期 |
| Table 2 | A1 | 修改後的值 | 描述 |
```

## 🔧 技術棧

| 技術 | 用途 |
|------|------|
| Python 3.x | 主要語言 |
| pandas | Excel 數據處理 |
| pywin32 | Word COM 操作 |
| openpyxl | Excel 讀寫 |
| win32com | Windows COM 接口 |

## 📁 文件結構

```
auto_report-table_modify_piero/
├── auto_report_table_modify.py     # 主程式
├── Master_Log.xlsx                 # 配置範例
├── Template.doc                    # Word 範本
├── Final_Report.docx               # 生成結果示例
└── README.md                       # 本文件
```

## ⚠️ 注意事項

1. 表格必須有明確的名稱或標識
2. Excel 配置的單元格位置格式要正確（如 B2、C3）
3. 執行前備份原始 Word 文件
4. Windows 系統專用

## 🐛 已知限制

- 僅支援 Windows 系統
- 需要 Microsoft Office 安裝
- 不支援跨越多個 Word 文件的批量操作

## 📝 更新日誌

### v1.0 (2024-01)
- 首個穩定版本
- 支援表格值替換和狀態更新

## 🤝 相關工具

- [auto_report_table_date.py](../auto_report-table-date_piero/) - 日期填入工具
- [integrated_report_tool](../integrated_report_tool/) - 整合工具
- [find_tables_in_word](../find_tables_in_word/) - 表格檢測工具

## 📞 技術支援

如有問題或建議，請聯絡開發者或提交 Issue。
