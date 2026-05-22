# 核心報告生成工具 - Piero 版本 (v6)

**版本**: v6  
**開發者**: Piero  
**狀態**: ✅ 生產版本

## 🎯 功能

用於自動化生成 Word RT 報告的核心工具。Piero 的穩定版本，支援動態路徑配置、自動章節管理和目錄更新。

## ✨ 主要特性

- **動態路徑設定**: 相對於腳本位置的動態路徑配置，便於跨機器使用
- **自動章節移除**: 根據 Excel 配置自動刪除跳過的章節
- **目錄自動更新**: 自動執行 Word 目錄更新
- **配置管理**: 提供 create_config.py 簡化配置創建
- **完整的章節邏輯**: 包含標題清理和編號更新
- **穩定可靠**: 經過多次迭代和測試

## 📋 使用方式

### 環境需求
- Python 3.7+
- pandas
- pywin32
- openpyxl
- win32com

### 快速開始

```python
# 方式 1: 使用現有配置
python auto_report.py

# 方式 2: 生成新配置
python create_config.py
# 然後編輯生成的配置文件
python auto_report.py
```

### 配置說明

**Excel 檔案格式** (Master_Log.xlsx):
```
| 章節名稱 | 測試狀態 | 其他數據 |
|---------|---------|--------|
| Module A | 完成 | ... |
| Module B | 跳過 | ... |
| Module C | 完成 | ... |
```

**配置文件** (config.py):
```python
EXCEL_PATH = './Master_Log.xlsx'
TEMPLATE_PATH = './Template.doc'
OUTPUT_PATH = './Final_Report.docx'
TARGET_SHEET = 'Sheet1'
STATUS_COLUMN = '測試狀態'
SKIP_KEYWORD = '跳過'
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
generate_RT_report_piero/
├── auto_report.py                  # 主程式
├── create_config.py                # 配置創建工具
├── Master_Log.xlsx                 # 配置範例
├── Template.doc                    # Word 範本
├── Final_Report.docx               # 生成結果示例
└── README.md                       # 本文件
```

## 🔍 主要函數

### auto_report.py 的核心功能

1. **config_load()**: 載入配置文件
2. **excel_read()**: 讀取 Excel 數據
3. **word_remove_sections()**: 移除跳過的章節
4. **word_update_toc()**: 更新目錄
5. **word_cleanup()**: 清理空標題
6. **word_renumber_tables()**: 重新編號表格

## ⚠️ 注意事項

1. 此工具基於 COM 接口，Windows 系統專用
2. 需要安裝 Microsoft Office
3. Excel 和 Word 文件在執行時應關閉
4. 執行前備份原始文件
5. 第一次執行可能較慢（COM 初始化）

## 🐛 已知限制

- Windows 系統專用
- 不支援 LibreOffice 或其他辦公軟件
- 複雜的 VBA 宏可能有衝突
- 非常大的 Word 文件可能性能較差

## 💡 最佳實踐

1. **配置驗證**: 執行前驗證 Excel 和 Word 文件路徑
2. **備份策略**: 保留原始範本副本
3. **測試運行**: 先用小樣本測試
4. **逐步執行**: 可在 auto_report.py 中註釋部分功能逐步測試

## 📝 更新日誌

### v6 (2024-01)
- 穩定版本
- 完整的章節管理
- 動態路徑支援

### v5 (2023-12)
- 新增目錄更新功能
- 改進錯誤處理

## 🤝 相關工具

- [generate_RT_report_luke](../generate_RT_report_luke/) - Luke 版本（Run 層級替換）
- [integrated_report_tool](../integrated_report_tool/) - 最新整合版
- [auto_report-table-date_piero](../auto_report-table-date_piero/) - 日期填入工具

## 📞 技術支援

如有問題或建議，請聯絡開發者或提交 Issue。
