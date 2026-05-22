# 核心報告生成工具 - Luke 版本

**版本**: v1.0  
**開發者**: Luke  
**狀態**: ✅ 生產版本

## 🎯 功能

用於自動化生成 Word RT 報告的核心工具。基於 Python-docx 的 Run 對象層級替換方案，實現格式保留的高質量內容替換。

## ✨ 主要特性

- **格式保留替換**: 在 Run 層級進行替換，完全保留原始格式（字體、顏色、粗體等）
- **無格式丟失**: 相比段落層級替換，避免格式被破壞
- **靈活配置**: 支援 Excel 驅動的動態配置
- **多文件支援**: 支援批量處理多個 Word 文件
- **詳細日誌**: 完整的執行過程記錄

## 📋 使用方式

### 環境需求
- Python 3.7+
- pandas
- python-docx
- openpyxl

### 快速開始

```python
# 1. 準備 Excel 配置文件 (data.xlsx)
#    包含要替換的字段和對應的值

# 2. 準備 Word 範本 (template.docx)

# 3. 執行腳本
python docx_replace.py
```

### 配置說明

**Excel 檔案格式**:
```
| 字段名稱 | 替換值 |
|---------|--------|
| {{TITLE}} | Test Report 2024 |
| {{DATE}} | 2024-01-15 |
| {{VERSION}} | v1.0 |
| {{ENGINEER}} | John Doe |
```

**Word 範本說明**:
- 在範本中使用 `{{FIELD_NAME}}` 作為佔位符
- 支援在任何位置使用佔位符（標題、段落、表格等）

## 🔧 技術棧

| 技術 | 用途 |
|------|------|
| Python 3.x | 主要語言 |
| python-docx | Word 文件操作（Run 層級） |
| pandas | Excel 配置處理 |
| openpyxl | Excel 讀寫 |

## 📁 文件結構

```
generate_RT_report_luke/
├── docx_replace.py                 # 主程式（核心算法）
├── requirements.txt                # 依賴包列表
├── data.xlsx                       # 配置範例
├── template.docx                   # Word 範本
├── run_method_result.docx          # 生成結果示例
└── README.md                       # 本文件
```

## 🔍 技術亮點

### Run 層級替換
```python
# 傳統做法（會丟失格式）
paragraph.text = new_text

# Luke 的做法（保留格式）
for run in paragraph.runs:
    if placeholder in run.text:
        run.text = run.text.replace(placeholder, new_value)
        # 格式自動保留！
```

## ⚠️ 注意事項

1. 佔位符必須使用 `{{FIELD_NAME}}` 格式
2. 佔位符不能跨越多個 Run（可在 Word 中檢查）
3. 執行前備份原始文件
4. 支援 .docx 格式（推薦）

## 🐛 已知限制

- 跨越多個 Run 的佔位符無法識別（需在 Word 中調整）
- 表格內的複雜嵌套可能有問題
- 不支援 .doc 舊格式

## 💡 最佳實踐

1. **佔位符設計**: 確保每個佔位符位於單個 Run 內
2. **測試替換**: 先用簡單數據測試，確保格式保留
3. **備份文件**: 始終保留原始範本副本
4. **結構化數據**: 在 Excel 中清晰組織配置數據

## 📝 更新日誌

### v1.0 (2024-01)
- 首個穩定版本
- Run 層級替換實現
- 格式完整保留

## 🤝 相關工具

- [generate_RT_report_piero](../generate_RT_report_piero/) - Piero 穩定版
- [integrated_report_tool](../integrated_report_tool/) - 整合工具
- [doc_to_template](../doc_to_template/) - 文件轉換

## 📞 技術支援

如有問題或建議，請聯絡開發者或提交 Issue。
