# Generate RT Report - Word 自動化報告生成工具集

[English](README.md) | [繁體中文](README.zh-TW.md)

完整的 Microsoft Word RT 報告自動化解決方案。包含報告生成、表格管理、大綱處理等多項功能，支援批量化、智慧化的報告生成流程。

## 🎯 核心功能

### 🌟 推薦工具

| 工具 | 說明 | 版本 | 狀態 |
|------|------|------|------|
| **report_generator** | 統一工具（Luke + Piero 整合），7 階段處理 | v1.0 | ✅ **推薦 — 請使用此工具** |

### 獨立工具 (穩定)

| 工具 | 說明 | 狀態 |
|------|------|------|
| **utils/date_filler** | 表格日期填入與章節管理 | ✅ 穩定 |
| **utils/outline_filter** | 大綱智慧過濾 | ✅ 最新 (v8.0) |

### 舊版報告生成器 (僅供參考)

已被 `report_generator` 取代，保留供參考。

| 工具 | 說明 | 版本 |
|------|------|------|
| **legacy/integrated_report** | 早期整合版 | v1.0 |
| **legacy/piero_v6** | 核心報告生成（v6） | v6 |
| **legacy/v33_claude** | AI 輔助實驗版（`final` 的程式碼基礎） | v33 |

### 輔助工具 (補充功能)

| 工具 | 說明 | 用途 |
|------|------|------|
| **utils/format_replace** | 格式保留替換邏輯 | Run 層級精細控制 |
| **utils/table_modifier** | 表格內容修改工具 | 修改表格數值和狀態 |
| **utils/table_finder** | 表格檢測與分析 | 表格位置和結構分析 |
| **utils/outline_reader** | 大綱讀取工具 | 文檔結構解析 |

## 📁 專案結構

```
generate_RT_report/
├── 🌟 report_generator/            # 統一報告生成器（推薦）
│   ├── main.py                      # 核心程式（7 階段處理）
│   ├── QUICKSTART.md                # 快速開始指南（從這裡開始）
│   ├── README.md                    # 完整文檔
│   └── EXAMPLES.md                  # 範例文件指南
│
├── utils/                           # 單一用途獨立工具
│   ├── date_filler/                 # 表格日期填入與章節管理
│   │   └── auto_report_table_date.py
│   ├── outline_filter/              # 大綱智慧過濾（v8.0）
│   │   └── word_outline_filter.py
│   ├── table_modifier/              # 表格內容修改
│   │   └── auto_report_table_modify.py
│   ├── table_finder/                # 表格檢測與分析
│   │   └── find_tables_in_word.py
│   ├── outline_reader/              # 大綱讀取 / 結構解析
│   │   └── read_word_outline.py
│   └── format_replace/              # Run 層級格式保留替換
│       ├── docx_replace.py
│       └── requirements.txt
│
├── legacy/                          # 已被 report_generator 取代 — 凍結，僅供參考
│   ├── integrated_report/           # 早期整合版（v1.0）
│   │   ├── unified_report.py
│   │   ├── create_config.py
│   │   └── Report_Config.xlsx
│   ├── piero_v6/                    # 核心報告生成（v6）
│   │   ├── auto_report.py
│   │   └── create_config.py
│   └── v33_claude/                  # AI 輔助實驗版（report_generator 的程式碼基礎）
│       └── auto_report_v33.py
│
├── README.md
└── README.zh-TW.md
```

## 🚀 快速開始

### 選項 A: 使用新統一工具 (推薦 - 從這裡開始!)

```bash
cd report_generator
cat QUICKSTART.md                 # 先讀指南
pip install python-docx openpyxl pywin32

# 準備檔案（自行準備）
# - Master_Log.xlsx  (填入測試數據)
# - Template.doc     (Word 報告範本)

python main.py                    # 輸出: Final_Report.docx
```

### 選項 B (舊版): 使用整合工具

> ⚠️ 已被 `report_generator` 取代 — 保留供參考。

```bash
cd legacy/integrated_report
python create_config.py           # 創建配置
python unified_report.py          # 執行報告生成
```

### 選項 C (舊版): 使用 Piero 版本

> ⚠️ 已被 `report_generator` 取代 — 保留供參考。

```bash
cd legacy/piero_v6
python auto_report.py             # 執行報告生成
```

### 選項 D: 使用日期填入工具

```bash
cd utils/date_filler
python auto_report_table_date.py  # 填入日期和時間
```

## 📋 版本對比與選擇指南

### 該使用哪個版本？

| 使用場景 | 推薦工具 | 理由 |
|---------|---------|------|
| 報告生成（任何場合） | **report_generator** | 唯一推薦入口，7 階段處理 |
| 只需填寫日期和時間 | utils/date_filler | 專用工具，功能精準 |
| 需要精細格式控制 | utils/format_replace | Run 層級替換，格式完全保留 |
| 需要分析表格結構 | utils/table_finder | 專用表格分析工具 |
| 需要過濾大綱內容 | utils/outline_filter | 最新版本（v8.0），功能強大 |

## 📄 Master_Log.xlsx 欄位規格 (report_generator)

`main.py` 讀取 `Master_Log.xlsx` 的**第一個工作表**（active sheet）；第 1 列必須為欄位標題。必要欄位：

| 欄位 | 說明 |
|------|------|
| `測項名稱` | 測試項目名稱。必須與範本中的標題文字一致（用於定位、保留或刪除章節） |
| `測試狀態` | 測試狀態，例如 `Pass`。值為 `跳過` 的列會連同範本中對應章節一併刪除 |
| `開始時間` | 開始時間 |
| `結束時間` | 結束時間 |
| `Physical Damage` | 物理損傷檢查結果 |
| `Functional Check` | 功能檢查結果 |

注意：
- 沒有 `測項名稱` 值的列會被忽略。
- 額外欄位（例如 `測項編號`、`測試結果`、`照片路徑`）允許存在，生成器不會使用。
- 範本（`Template.doc`）應包含佔位符 `{Report_Title}`、`{Report_Date}`、`{Author}`、**3.1 "Test Item Result"** 表格（4 欄以上：名稱 / 開始 / 結束 / 狀態），以及含 "Physical Damage" / "Functional Check" 標題的 4.x 詳細結果表格。

## 🔄 工作流程示例

### 完整自動化流程

```
1. 準備 Excel 配置 (Master_Log.xlsx 或 Report_Config.xlsx)
   ↓
2. 準備 Word 範本 (Template.doc 或 Template.docx)
   ↓
3. 執行報告生成
   - 使用 report_generator (推薦)
   - 舊版替代：legacy/integrated_report / legacy/piero_v6
   ↓
4. （可選）使用 utils/date_filler 填入日期
   ↓
5. （可選）使用 utils/outline_filter 過濾內容
   ↓
6. 最終報告 Final_Report.docx
```

## 📚 詳細文檔

每個工具的詳細說明請參考各資料夾的 README.md：

* [legacy/integrated_report/README.md](legacy/integrated_report/README.md) - 整合工具完整文檔
* [legacy/piero_v6/README.md](legacy/piero_v6/README.md) - Piero 版本說明
* [utils/format_replace/README.md](utils/format_replace/README.md) - Luke 版本說明
* [utils/date_filler/README.md](utils/date_filler/README.md) - 日期工具說明
* [utils/outline_filter/README.md](utils/outline_filter/README.md) - 大綱過濾工具說明
* [其他工具文檔](.) - 請查看各資料夾

## ⚙️ 環境需求

### 通用需求
- **作業系統**: Windows (大多數工具基於 COM 接口)
- **Python**: 3.7 或更高版本
- **Microsoft Office**: Word 和 Excel (用於 COM 操作)

### Python 依賴包
```bash
pip install pandas pywin32 openpyxl python-docx win32com
```

### 首次配置
```bash
# 完成 pywin32 配置（僅首次需要）
python -m pip install --upgrade pywin32
python Scripts/pywin32_postinstall.py -install
```

## ⚠️ 重要注意事項

1. **備份文件**: 執行任何自動化操作前，務必備份原始 Word 和 Excel 文件
2. **關閉文件**: 執行時請關閉所有相關的 Word 和 Excel 文件，避免權限衝突
3. **配置驗證**: 執行前驗證 Excel 配置文件的欄位名稱和路徑
4. **首次執行**: 首次執行可能較慢（COM 初始化），請耐心等待

## 🔧 技術棧

| 技術 | 用途 |
|------|------|
| Python 3.x | 主要開發語言 |
| pandas | 數據處理 |
| python-docx | Word 文件操作 |
| pywin32 | Windows COM 接口 |
| openpyxl | Excel 操作 |

## 📝 版本歷史與整理

### 工具收斂與欄位規格 (2026-10)
- ✅ `report_generator` 為唯一推薦的報告生成工具
- ✅ `legacy/integrated_report`、`legacy/piero_v6`、`legacy/v33_claude` 標記為**舊版**（僅供參考）
- ✅ 新增 `Master_Log.xlsx` 欄位規格（見上方）

### 隱私淨化 (2026-10)
- ✅ 移除內部報告範本與範例數據文件（含內部專案資訊）
- ✅ 輸入文件（`Master_Log.xlsx`、`Template.doc`）需由使用者自行提供

### 本次整理 (2024-01)
- ✅ 刪除廢棄文件和重複版本
- ✅ 重命名文件規範化
- ✅ 為所有工具補充詳細 README
- ✅ 統一版本標記
- ✅ 更新根目錄文檔
- ✅ 整理重複代碼

**整理結果**：
- 從 15+ 個資料夾精簡至 10 個有效資料夾
- 刪除 27 個廢棄文件
- 刪除 12 個舊版本輸出
- 為 9 個資料夾補充詳細文檔

## 🤝 貢獻者

- **Luke**: 格式保留替換邏輯、大綱過濾工具優化
- **Piero**: 核心報告生成、日期填入工具、表格管理
- **Claude AI**: 新版本實驗、整合優化、文檔完善

## 📞 技術支援

如有問題或建議：
1. 查看各工具的 README.md 文檔
2. 檢查 Excel 配置文件格式
3. 驗證 Word 和 Excel 文件是否正確關閉
4. 聯絡相應工具的開發者

## 📄 許可證與使用

本專案內容供內部使用。使用前請確保了解各工具的要求和限制。

---

**推薦工作流程**: 
1. 使用 `report_generator`（唯一推薦入口）
2. 按需搭配獨立工具（日期填入、大綱過濾）
3. 舊版生成器（`legacy/integrated_report`、`legacy/piero_v6`）僅供參考
