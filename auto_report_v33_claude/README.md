# 自動化報告生成工具 v33 (Claude 版本)

**版本**: v33  
**開發者**: Claude AI  
**狀態**: 🔬 實驗版本

## 🎯 功能

Advanced Word RT 報告自動化工具，整合了 AI 智慧功能，可自動識別和提取報告的基本信息（標題、編號、版本等），並進行動態的內容替換和格式調整。

## ✨ 主要特性

- **智慧信息識別**: 從 Excel 第一列自動識別報告基本資訊
- **自動提取標題**: 智慧判斷並提取報告名稱、版本號、編號
- **動態內容替換**: 支援多層級的模板替換
- **格式智慧調整**: 自動適應不同的 Word 文件格式
- **詳細日誌**: 完整的執行過程記錄，便於調試

## 📋 使用方式

### 環境需求
- Python 3.7+
- pandas
- pywin32
- openpyxl
- win32com

### 快速開始

```python
# 1. 準備 Excel 配置文件
#    第一列包含報告基本信息：
#    - 報告名稱
#    - 報告編號
#    - 版本號
#    - 其他可選信息

# 2. 準備 Word 範本

# 3. 執行腳本
python auto_report_v33.py
```

### 配置說明

**Excel 檔案格式**:
```
| 項目 | 數值 |
|------|------|
| 報告名稱 | Test Report 2024 |
| 報告編號 | RD-2024-001 |
| 版本號 | v33 |
| 測試日期 | 2024-01-15 |
| 測試工程師 | John Doe |
```

## 🔧 技術棧

| 技術 | 用途 |
|------|------|
| Python 3.x | 主要語言 |
| pandas | Excel 數據處理和智慧識別 |
| pywin32 | Word COM 操作 |
| openpyxl | Excel 讀寫 |
| win32com | Windows COM 接口 |

## 📁 文件結構

```
auto_report_v33_claude/
├── auto_report_v33.py              # 主程式（Claude 開發）
├── Master_Log.xlsx                 # 配置範例
├── Template.doc                    # Word 範本
├── README.txt                      # 簡要說明
└── README.md                       # 本文件
```

## ⚠️ 注意事項

1. 此版本為實驗版本，可能存在不穩定性
2. Excel 配置必須包含完整的報告基本信息
3. 執行前務必備份原始文件
4. Windows 系統專用

## 🔬 實驗特性

本版本包含以下實驗特性，可能在未來版本中調整：
- AI 輔助的信息識別
- 自動格式檢測和調整
- 擴展的 Excel 配置支援

## 🐛 已知問題

- 某些特殊格式的 Word 文件可能無法正確處理
- 大型 Excel 文件處理可能較慢
- AI 識別在特殊情況下可能有誤差

## 📝 更新日誌

### v33 (2024-01)
- Claude AI 參與開發
- 新增智慧信息識別功能
- 改進格式適應能力

## 🤝 相關工具

- [integrated_report_tool](../integrated_report_tool/) - 整合生產版工具
- [auto_report-table-date_piero](../auto_report-table-date_piero/) - 穩定版日期填入
- [generate_RT_report_piero](../generate_RT_report_piero/) - 穩定版報告生成

## 💡 反饋與建議

此為實驗版本，歡迎提交反饋和建議，幫助改進工具功能。

## 📞 技術支援

如有問題，請聯絡開發者或查閱相關文檔。
