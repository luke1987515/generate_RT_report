# Word 大綱過濾工具 (Word Outline Filter)

**版本**: v8.0 (Luke 最新版)  
**開發者**: Luke（多次迭代）  
**狀態**: ✅ 生產版本

## 🎯 功能

用於對 Word 文件進行智慧大綱過濾的工具。根據特定條件（如關鍵詞、標題等級等）過濾 Word 文件的大綱內容，生成精簡版本。

## ✨ 主要特性

- **條件過濾**: 根據關鍵詞、標題等級等條件進行過濾
- **批量處理**: 支援批量處理多個 Word 文件
- **多個版本**: 支援生成多個不同的過濾版本（v2 → v8）
- **結構保留**: 保留 Word 文件的層級結構和格式
- **Excel 配置**: 支援 Excel 配置驅動的過濾規則
- **詳細日誌**: 記錄每個過濾步驟

## 📋 使用方式

### 環境需求
- Python 3.7+
- python-docx
- pandas
- openpyxl

### 快速開始

```python
# 方式 1: 基本使用
python word_outline_filter.py Template.doc

# 方式 2: 使用配置文件
python word_outline_filter.py Template.doc --config outline.xlsx

# 方式 3: 指定過濾規則
python word_outline_filter.py Template.doc --keywords "Module,Test" --level 2
```

### 配置說明

**outline.xlsx** 配置文件：

**Sheet: 過濾規則**
```
| 規則 ID | 類型 | 值 | 操作 | 優先級 |
|--------|------|------|------|------|
| rule_1 | 關鍵詞 | Module | 包含 | 1 |
| rule_2 | 等級 | 2 | 保留 | 2 |
| rule_3 | 關鍵詞 | Skip | 移除 | 3 |
```

**Sheet: 版本設置**
```
| 版本 | 描述 | 啟用規則 |
|------|------|---------|
| Standard | 標準版本 | rule_1, rule_2 |
| Minimal | 最小版本 | rule_2 |
| Full | 完整版本 | rule_1 |
```

## 🔧 技術棧

| 技術 | 用途 |
|------|------|
| Python 3.x | 主要語言 |
| python-docx | Word 文件操作 |
| pandas | Excel 配置處理 |
| openpyxl | Excel 讀寫 |

## 📁 文件結構

```
word_outline_filter/
├── word_outline_filter.py          # 主程式
├── outline.xlsx                    # 過濾規則配置
├── outline_new.xlsx                # 新的配置版本
├── test_v3.xlsx                    # v3 版本測試配置
├── requirements.txt                # 依賴包列表
├── Template.doc                    # 原始 Word 範本
├── Filtered_Template_luke_v8.docx  # Luke v8 最新版本（推薦）
└── README.md                       # 本文件
```

## 📝 版本演進

| 版本 | 改進 | 特點 |
|------|------|------|
| v2 | 初始版本 | 基本過濾功能 |
| v3 | 改進過濾邏輯 | 支援多規則組合 |
| v4 | 優化性能 | 更快的大文件處理 |
| v5 | 增強靈活性 | 自訂過濾函數支援 |
| v6 | 改進配置 | Excel 配置驅動 |
| v7 | 批量處理 | 支援多文件處理 |
| v8 | 最新版本 | 完整功能 + 性能優化 |

**推薦使用 v8.0 最新版本**

## 🔍 過濾規則詳解

### 關鍵詞過濾
```python
# 匹配包含"Module"的標題
keywords=['Module']
operation='contain'

# 匹配不包含"Skip"的標題
keywords=['Skip']
operation='exclude'
```

### 等級過濾
```python
# 保留等級 <= 2 的標題
level=2
operation='level_le'

# 保留等級 == 1 的標題
level=1
operation='level_eq'
```

### 複合過濾
```python
# (keyword='Module' AND level<=2) OR keyword='Summary'
rules = [
    {'keyword': 'Module', 'level': 2, 'op': 'and'},
    {'keyword': 'Summary', 'op': 'or'}
]
```

## ⚠️ 注意事項

1. 過濾規則應該清晰明確，避免誤刪重要內容
2. 首次過濾應驗證結果
3. 保留原始文件備份
4. 複雜規則應分步測試

## 🐛 已知限制

- 嵌套很深的結構可能處理較慢
- 複雜的 VBA 宏不支援
- 某些特殊格式可能無法完全保留

## 💡 最佳實踐

1. **規則設計**: 先設計簡單規則，逐步增加複雜性
2. **版本管理**: 為不同用途創建不同版本
3. **測試運行**: 在小文件上先測試規則
4. **檔案保留**: 保留所有版本以便比較

## 📝 更新日誌

### v8.0 (2024-01) - 最新版
- Luke 的最新迭代
- 完整的功能支援
- 性能最優化
- 推薦用於生產環境

### v7.0
- 增加批量處理
- 改進日誌記錄

### v6.0
- Excel 配置驅動
- 多規則支援

## 🌟 推薦配置

對於大多數使用場景：

```python
# 推薦配置
filter_rules = {
    'level': 2,  # 保留 H1, H2 層級
    'keywords': ['Test', 'Module', 'Result'],  # 包含這些關鍵詞
    'exclude': ['Draft', 'TODO', 'Internal'],  # 排除這些
    'preserve_numbering': True,  # 保留編號
    'keep_structure': True  # 保留層級結構
}
```

## 🤝 相關工具

- [read_word_outline](../read_word_outline/) - 大綱讀取工具
- [find_tables_in_word](../find_tables_in_word/) - 表格檢測
- [doc_to_template](../doc_to_template/) - 文件轉換

## 📞 技術支援

如有問題或建議，請聯絡開發者或提交 Issue。

---

**建議**: 使用最新版本 v8.0 以獲得最佳性能和功能支援。
