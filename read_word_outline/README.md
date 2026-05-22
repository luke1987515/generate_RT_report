# Word 大綱讀取工具 (Read Word Outline)

**版本**: v1.0  
**開發者**: 項目團隊  
**狀態**: ✅ 生產版本

## 🎯 功能

用於讀取和解析 Word 文件大綱結構的工具。能夠提取文檔的層級結構，包括標題、小節和內容組織方式。

## ✨ 主要特性

- **大綱提取**: 完整提取 Word 文件的大綱結構
- **層級分析**: 識別標題的層級關係（H1, H2, H3 等）
- **結構映射**: 生成清晰的文檔結構樹
- **位置追蹤**: 記錄每個標題在文件中的位置
- **Excel 匯出**: 可將大綱匯出為 Excel 格式
- **嵌套支援**: 支援複雜的多層嵌套結構

## 📋 使用方式

### 環境需求
- Python 3.7+
- python-docx

### 快速開始

```python
# 方式 1: 命令行使用
python read_word_outline.py Template.doc

# 方式 2: Python 代碼使用
from read_word_outline import OutlineReader

reader = OutlineReader('Template.doc')
outline = reader.read_outline()
for item in outline:
    print(f"{'  ' * item['level']}{item['title']}")
```

### 輸出說明

工具會輸出清晰的大綱結構：

```
1. 簡介
   1.1 背景
   1.2 目的
2. 測試方法
   2.1 環境配置
   2.2 測試流程
3. 結果分析
   3.1 性能測試
   3.2 功能驗證
4. 結論
```

## 🔧 技術棧

| 技術 | 用途 |
|------|------|
| Python 3.x | 主要語言 |
| python-docx | Word 文件解析 |
| openpyxl | Excel 匯出 |

## 📁 文件結構

```
read_word_outline/
├── read_word_outline.py            # 主程式
├── Template.doc                    # 範例 doc 文件
└── README.md                       # 本文件
```

## 📝 詳細說明

### 大綱等級定義

| 等級 | 描述 | 對應 Word 樣式 |
|------|------|----------------|
| 0 | 標題 1 (H1) | Heading 1 |
| 1 | 標題 2 (H2) | Heading 2 |
| 2 | 標題 3 (H3) | Heading 3 |
| 3+ | 標題 4+ (H4+) | Heading 4+ |

### 提取的信息

- 標題文本
- 標題等級
- 在文件中的位置（段落編號）
- 頁碼（如可取得）

## ⚠️ 注意事項

1. 確保 Word 文件使用正確的大綱樣式（Heading 1, 2, 3 等）
2. 自定義樣式可能無法識別
3. 支援 .doc 和 .docx 格式

## 🐛 已知限制

- 某些自定義樣式可能無法識別
- 嵌入的圖表或物件不會被解析
- 非常大的文件可能解析較慢

## 💡 使用場景

1. **文檔分析**: 快速瞭解 Word 文件的結構
2. **內容重組**: 基於大綱進行內容重組或提取
3. **質量檢查**: 檢查文檔的層級結構是否合理
4. **批量處理**: 批量讀取多個文件的大綱

## 📝 更新日誌

### v1.0 (2024-01)
- 首個穩定版本
- 支援基本大綱提取
- Excel 匯出功能

## 🤝 相關工具

- [word_outline_filter](../word_outline_filter/) - 大綱過濾工具
- [find_tables_in_word](../find_tables_in_word/) - 表格檢測
- [doc_to_template](../doc_to_template/) - 文件轉換

## 📞 技術支援

如有問題或建議，請聯絡開發者或提交 Issue。
