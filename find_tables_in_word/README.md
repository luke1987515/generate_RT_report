# 表格檢測工具 (Find Tables in Word)

**版本**: v1.0  
**開發者**: 項目團隊  
**狀態**: ✅ 生產版本

## 🎯 功能

用於在 Word 文件中定位、提取和分析表格的工具。支援複雜的嵌套表格、表格結構分析和詳細的位置信息提取。

## ✨ 主要特性

- **表格定位**: 快速定位 Word 文件中所有表格
- **結構分析**: 詳細分析表格結構（行數、列數、單元格內容）
- **嵌套支援**: 支援檢測和處理嵌套表格
- **位置提取**: 提取表格在文件中的位置信息
- **標題關聯**: 自動識別表格標題和描述
- **Excel 匯出**: 將表格信息匯出為 Excel 格式便於分析

## 📋 使用方式

### 環境需求
- Python 3.7+
- python-docx

### 快速開始

```python
# 方式 1: 命令行使用
python find_tables_in_word.py Template.doc

# 方式 2: Python 代碼使用
from find_tables_in_word import TableFinder

finder = TableFinder('Template.doc')
tables = finder.find_all_tables()
for table_info in tables:
    print(f"找到表格: {table_info['position']}")
    print(f"大小: {table_info['rows']} 行 x {table_info['cols']} 列")
```

### 輸出說明

工具會輸出以下信息：

```
表格 #1:
  位置: 第 2 段落
  大小: 5 行 x 3 列
  標題: Test Results Summary
  內容: [表格內容詳情]

表格 #2:
  位置: 第 15 段落
  大小: 10 行 x 4 列
  標題: Module Test Status
  內容: [表格內容詳情]
```

## 🔧 技術棧

| 技術 | 用途 |
|------|------|
| Python 3.x | 主要語言 |
| python-docx | Word 文件解析 |
| pandas | 數據分析和匯出 |

## 📁 文件結構

```
find_tables_in_word/
├── find_tables_in_word.py          # 主程式
├── Template.doc                    # 範例 doc 文件
├── Template.docx                   # 範例 docx 文件
├── README.md                       # 本文件
└── README.txt                      # 簡要說明
```

## 📝 詳細說明

### 表格信息提取

工具可以提取以下信息：

1. **基本信息**
   - 表格位置（段落編號）
   - 行列數量
   - 單元格內容

2. **結構信息**
   - 是否有合併單元格
   - 嵌套情況
   - 表格邊框和樣式

3. **內容信息**
   - 文本內容
   - 格式資訊（粗體、斜體等）
   - 超連結

## ⚠️ 注意事項

1. 支援 .doc 和 .docx 格式
2. 複雜的嵌套表格可能需要逐層分析
3. 某些舊版 .doc 文件可能解析不完整

## 🐛 已知限制

- 複雜的 VBA 宏不支援
- 某些特殊的表格格式可能無法識別
- 非常大的表格（>1000 行）可能影響性能

## 💡 使用場景

1. **表格審計**: 快速檢查文件中所有表格
2. **數據提取**: 從 Word 報告提取表格數據
3. **結構分析**: 分析複雜文件的結構
4. **格式驗證**: 驗證表格是否符合預期格式

## 📝 更新日誌

### v1.0 (2024-01)
- 首個穩定版本
- 支援基本表格檢測
- Excel 匯出功能

## 🤝 相關工具

- [doc_to_template](../doc_to_template/) - 文件轉換
- [word_outline_filter](../word_outline_filter/) - 大綱過濾
- [read_word_outline](../read_word_outline/) - 大綱讀取

## 📞 技術支援

如有問題或建議，請聯絡開發者或提交 Issue。
