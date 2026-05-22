# 文件與範本轉換工具 (Doc to Template)

**版本**: v1.0  
**開發者**: 項目團隊  
**狀態**: ✅ 生產版本

## 🎯 功能

用於在 Word 文檔（.doc）和範本格式（.docx）之間進行雙向轉換的工具。支援批量轉換和智慧檔案格式識別。

## ✨ 主要特性

- **雙向轉換**: 支援 doc ↔ docx ↔ template 的相互轉換
- **批量處理**: 可一次性轉換多個文件
- **格式保留**: 轉換過程中保留原始格式和樣式
- **智慧識別**: 自動識別檔案類型，選擇合適的轉換方式
- **範本管理**: 支援從文件生成範本，也支援從範本還原文件

## 📋 使用方式

### 環境需求
- Python 3.7+
- python-docx
- openpyxl

### 快速開始

```python
# 方式 1: 文檔轉範本
python doc_to_template.py input.doc output_template.docx

# 方式 2: 範本轉文檔
python template_to_doc.py input_template.docx output.doc

# 方式 3: 批量轉換
python doc_to_template.py --batch ./documents/ ./templates/
```

### 詳細用法

```python
from doc_to_template import DocConverter

# 初始化轉換器
converter = DocConverter()

# 轉換單個文件
converter.doc_to_template('input.doc', 'output.docx')

# 批量轉換
converter.batch_convert('./source/', './destination/')

# 使用自訂配置
converter.doc_to_template(
    'input.doc', 
    'output.docx',
    preserve_styles=True,
    keep_formatting=True
)
```

## 🔧 技術棧

| 技術 | 用途 |
|------|------|
| Python 3.x | 主要語言 |
| python-docx | Word 文件操作 |
| openpyxl | 相關 Excel 支援 |

## 📁 文件結構

```
doc_to_template/
├── doc_to_template.py              # Doc 到 Template 轉換
├── template_to_doc.py              # Template 到 Doc 轉換
├── RD260309A09*.DOC                # 範例 doc 文件
├── RD260309A09*.docx               # 範例 docx 文件
├── RD260309A09*_Template.docx      # 範例 template 文件
├── RD260309A09*_new_values.docx    # 轉換後的文件
├── data.xlsx                       # 配置數據
└── README.md                       # 本文件
```

## 📝 轉換規則

### Doc → Template 轉換
1. 讀取 .doc 文件的內容和格式
2. 保留所有樣式和格式設定
3. 轉換為 .docx 範本格式
4. 標記可變部分為佔位符

### Template → Doc 轉換
1. 讀取範本中的佔位符
2. 替換為實際數據（來自 Excel）
3. 生成最終的 .doc 文件

## ⚠️ 注意事項

1. .doc 格式為舊格式，某些高級格式可能不支援
2. 轉換過程中複雜的 VBA 宏可能會丟失
3. 圖片和嵌入物件需要特殊處理
4. 建議在轉換前備份原始文件

## 🐛 已知限制

- 不支援 VBA 宏
- 複雜的表格嵌套可能有問題
- 某些特殊字體可能無法保留

## 💡 最佳實踐

1. **備份原始文件**: 轉換前始終備份源文件
2. **驗證轉換結果**: 轉換後檢查檔案是否完整
3. **使用標準格式**: 使用標準的 Word 格式，避免特殊擴展
4. **批量測試**: 批量轉換前先測試小樣本

## 📝 更新日誌

### v1.0 (2024-01)
- 首個穩定版本
- 支援雙向轉換
- 批量處理能力

## 🤝 相關工具

- [find_tables_in_word](../find_tables_in_word/) - 表格檢測
- [read_word_outline](../read_word_outline/) - 大綱讀取
- [word_outline_filter](../word_outline_filter/) - 大綱過濾

## 📞 技術支援

如有問題或建議，請聯絡開發者或提交 Issue。
