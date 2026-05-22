# 整合式報告生成工具 (Unified Report Tool)

**版本**: v1.0 整合版  
**開發者**: 項目團隊（整合 Luke + Piero 的工作）  
**狀態**: ⭐ **推薦生產版本**

## 🎯 功能

最新的整合式 Word RT 報告自動化工具。結合了 Luke 的格式保留替換邏輯和 Piero 的章節管理能力，提供全方位的報告自動化解決方案。

## ✨ 主要特性

- **整合最佳實踐**: 結合了多個版本的優點
- **格式完整保留**: 支援 Run 層級的精細替換
- **完整的章節管理**: 自動刪除、編號、目錄更新
- **智慧路徑設定**: 動態相對路徑，支援跨機器執行
- **Excel 配置驅動**: 支援多 Sheet 配置
- **自動範本尋找**: 自動檢測 Template.doc 或 Template.docx
- **錯誤恢復**: 完整的錯誤處理和日誌記錄
- **配置管理工具**: 提供 create_config.py 簡化配置

## 📋 使用方式

### 環境需求
- Python 3.7+
- pandas
- python-docx / pywin32
- openpyxl
- win32com

### 快速開始

```python
# 方式 1: 使用現有配置
python unified_report.py

# 方式 2: 生成配置
python create_config.py
python unified_report.py

# 方式 3: 指定配置文件
python unified_report.py --config Report_Config.xlsx
```

### 配置說明

**Report_Config.xlsx** 支援多 Sheet：

**Sheet: 基本配置**
```
| 項目 | 值 |
|------|------|
| Excel Data Path | Master_Log.xlsx |
| Word Template | Template.doc |
| Output Path | Final_Report_Unified.docx |
```

**Sheet: 替換規則**
```
| 佔位符 | 替換值 | 所屬Sheet | 欄位 |
|--------|--------|---------|------|
| {{TITLE}} | Test Report 2024 | 基本 | - |
| {{MODULE}} | Module A | Sheet1 | 模組名稱 |
```

**Sheet: 章節管理**
```
| 章節名稱 | 來源欄位 | 刪除條件 | 優先級 |
|---------|---------|--------|------|
| 簡介 | - | - | 1 |
| Module A | Sheet1 | 跳過 | 2 |
```

## 🔧 技術棧

| 技術 | 用途 |
|------|------|
| Python 3.x | 主要語言 |
| python-docx | Word 操作（推薦） |
| pywin32 | Word COM（備選） |
| pandas | Excel 處理和配置驅動 |
| openpyxl | Excel 讀寫 |
| win32com | Windows 接口 |

## 📁 文件結構

```
integrated_report_tool/
├── unified_report.py               # 主程式（整合邏輯）
├── create_config.py                # 配置創建工具
├── Report_Config.xlsx              # Excel 多 Sheet 配置
├── Template.doc                    # Word 範本
├── Final_Report_Unified.docx       # 生成結果示例
└── README.md                       # 本文件
```

## 🔍 整合的功能對比

| 功能 | Luke版 | Piero版 | 整合版 |
|------|--------|---------|--------|
| 格式保留替換 | ✅ | ⭕ | ✅ |
| 章節刪除管理 | ⭕ | ✅ | ✅ |
| 目錄自動更新 | ⭕ | ✅ | ✅ |
| 動態路徑配置 | ⭕ | ✅ | ✅ |
| 多 Sheet 支援 | ⭕ | ⭕ | ✅ |
| 自動路徑尋找 | ⭕ | ⭕ | ✅ |
| 完整日誌記錄 | ✅ | ✅ | ✅ |

## ⚠️ 注意事項

1. **選擇操作模式**：
   - 使用 python-docx：無需 Office，跨平台，但格式支援有限
   - 使用 pywin32：需要 Windows + Office，功能完整

2. **配置驗證**：執行前驗證所有路徑和欄位名稱

3. **備份策略**：執行前備份原始文件

4. **首次運行**：首次運行可能較慢（依賴初始化）

## 🐛 已知限制

- COM 模式（pywin32）仅支援 Windows 系統
- 非常大的 Word 文件可能影響性能
- 複雜的 VBA 宏可能有衝突

## 💡 最佳實踐

1. **配置測試**: 先用簡單配置測試功能
2. **逐步增加**: 逐步增加複雜配置
3. **版本控制**: 保留配置的版本歷史
4. **文檔維護**: 定期更新配置說明文檔

## 📝 更新日誌

### v1.0 整合版 (2024-01)
- 整合 Luke 和 Piero 的優點
- 新增多 Sheet 配置支援
- 新增自動路徑尋找
- 完整的錯誤處理

## 🌟 推薦使用場景

1. ✅ 需要完整自動化報告生成
2. ✅ 複雜的多階段報告流程
3. ✅ 需要跨機器和團隊協作
4. ✅ 長期專案維護

## 🤝 相關工具

- [generate_RT_report_luke](../generate_RT_report_luke/) - 格式保留參考
- [generate_RT_report_piero](../generate_RT_report_piero/) - 章節管理參考
- [auto_report-table-date_piero](../auto_report-table-date_piero/) - 日期填入補充工具
- [auto_report_v33_claude](../auto_report_v33_claude/) - 最新 AI 實驗版

## 📞 技術支援

如有問題或建議，請聯絡開發者或提交 Issue。

---

**💡 推薦**: 新項目優先使用本整合工具，它提供了最完整的功能和最佳的維護性。
