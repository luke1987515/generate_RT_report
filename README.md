# Generate RT Report - Complete Word Automation Solution

[English](README.md) | [繁體中文](README.zh-TW.md)

A complete automated Microsoft Word RT report generation solution. Includes report generation, table management, outline processing and more, with support for batch and intelligent report generation workflows.

## 🎯 Core Features

### 🌟 Recommended Tool

| Tool | Description | Version | Status |
|------|-------------|---------|--------|
| **report_generator** | Unified tool (Luke + Piero integration), 7-stage processing | v1.0 | ✅ **RECOMMENDED — use this** |

### Standalone Utilities (Stable)

| Tool | Description | Status |
|------|-------------|--------|
| **utils/date_filler** | Table date filling & section management | ✅ Stable |
| **utils/outline_filter** | Intelligent outline filtering | ✅ Latest (v8.0) |

### Legacy Generators (Reference Only)

Superseded by `report_generator`; kept for reference.

| Tool | Description | Version |
|------|-------------|---------|
| **legacy/integrated_report** | Earlier integrated version | v1.0 |
| **legacy/piero_v6** | Core report generation (v6) | v6 |
| **legacy/v33_claude** | AI-assisted experimental version (code base of `final`) | v33 |

### Supplementary Tools

| Tool | Description | Purpose |
|------|-------------|---------|
| **utils/format_replace** | Format-preserving replacement logic | Run-level fine control |
| **utils/table_modifier** | Table content modification tool | Modify table values and status |
| **utils/table_finder** | Table detection and analysis | Table location and structure analysis |
| **utils/outline_reader** | Outline reading tool | Document structure parsing |

## 📁 Project Structure

```
generate_RT_report/
├── 🌟 report_generator/            # Unified report generator (RECOMMENDED)
│   ├── main.py                      # Core program (7-stage processing)
│   ├── QUICKSTART.md                # Quick start guide (START HERE)
│   ├── README.md                    # Complete documentation
│   └── EXAMPLES.md                  # Example files guide
│
├── utils/                           # Standalone single-purpose tools
│   ├── date_filler/                 # Table date filling & section management
│   │   └── auto_report_table_date.py
│   ├── outline_filter/              # Intelligent outline filtering (v8.0)
│   │   └── word_outline_filter.py
│   ├── table_modifier/              # Table content modification
│   │   └── auto_report_table_modify.py
│   ├── table_finder/                # Table detection and analysis
│   │   └── find_tables_in_word.py
│   ├── outline_reader/              # Outline reading / structure parsing
│   │   └── read_word_outline.py
│   └── format_replace/              # Run-level format-preserving replacement
│       ├── docx_replace.py
│       └── requirements.txt
│
├── legacy/                          # Superseded by report_generator — frozen, reference only
│   ├── integrated_report/           # Earlier integrated version (v1.0)
│   │   ├── unified_report.py
│   │   ├── create_config.py
│   │   └── Report_Config.xlsx
│   ├── piero_v6/                    # Core report generation (v6)
│   │   ├── auto_report.py
│   │   └── create_config.py
│   └── v33_claude/                  # AI-assisted experimental (code base of report_generator)
│       └── auto_report_v33.py
│
├── README.md
└── README.zh-TW.md
```

## 🚀 Quick Start

### 🌟 Option A: Using New Unified Tool (Recommended - START HERE!)

```bash
cd report_generator
# Read the guide first
cat QUICKSTART.md           

# Install dependencies
pip install python-docx openpyxl pywin32

# Prepare files
# - Master_Log.xlsx  (fill with test data)
# - Template.doc     (Word report template)

# Generate report
python main.py

# Output: Final_Report.docx + report_generation.log
```

### Option B (Legacy): Using Integrated Tool

> ⚠️ Superseded by `report_generator` — kept for reference.

```bash
cd legacy/integrated_report
python create_config.py           # Create configuration
python unified_report.py          # Generate report
```

### Option C (Legacy): Using Piero Version

> ⚠️ Superseded by `report_generator` — kept for reference.

```bash
cd legacy/piero_v6
python auto_report.py             # Generate report
```

### Option D: Using Date Filling Tool

```bash
cd utils/date_filler
python auto_report_table_date.py  # Fill dates and times
```

## 📋 Version Comparison & Selection Guide

### Which version to use?

| Use Case | Recommended Tool | Reason |
|----------|------------------|--------|
| Report generation (any case) | **report_generator** | Single recommended entry point, 7-stage processing |
| Only need date/time filling | utils/date_filler | Dedicated tool, precise functionality |
| Need fine format control | utils/format_replace | Run-level replacement, complete format preservation |
| Need to analyze table structure | utils/table_finder | Dedicated table analysis tool |
| Need to filter outline content | utils/outline_filter | Latest version (v8.0), powerful features |

## 📄 Master_Log.xlsx Field Specification (report_generator)

`main.py` reads the **active sheet** of `Master_Log.xlsx`; row 1 must be the column headers. Required columns:

| Column | Description |
|--------|-------------|
| `測項名稱` | Test item name. Must match the heading text in the template (used to locate and keep/delete chapters) |
| `測試狀態` | Test status, e.g. `Pass`. Rows with `跳過` have their template chapters removed |
| `開始時間` | Start time |
| `結束時間` | End time |
| `Physical Damage` | Physical damage check result |
| `Functional Check` | Functional check result |

Notes:
- Rows without a `測項名稱` value are ignored.
- Extra columns (e.g. `測項編號`, `測試結果`, `照片路徑`) are allowed and simply not used by the generator.
- The template (`Template.doc`) should contain the placeholders `{Report_Title}`, `{Report_Date}`, `{Author}`, a **3.1 "Test Item Result"** table (4+ columns: name / start / end / status), and 4.x detail tables with "Physical Damage" / "Functional Check" headers.

## 🔄 Workflow Example

### Complete Automation Workflow

```
1. Prepare Excel configuration (Master_Log.xlsx or Report_Config.xlsx)
   ↓
2. Prepare Word template (Template.doc or Template.docx)
   ↓
3. Run report generation
   - Use report_generator (recommended)
   - Legacy alternatives: legacy/integrated_report / legacy/piero_v6
   ↓
4. (Optional) Use utils/date_filler to fill dates
   ↓
5. (Optional) Use utils/outline_filter to filter content
   ↓
6. Final report: Final_Report.docx
```

## 📚 Detailed Documentation

Detailed documentation for each tool can be found in the README.md files in their respective folders:

* [legacy/integrated_report/README.md](legacy/integrated_report/README.md) - Integrated tool complete documentation
* [legacy/piero_v6/README.md](legacy/piero_v6/README.md) - Piero version documentation
* [utils/format_replace/README.md](utils/format_replace/README.md) - Luke version documentation
* [utils/date_filler/README.md](utils/date_filler/README.md) - Date tool documentation
* [utils/outline_filter/README.md](utils/outline_filter/README.md) - Outline filter tool documentation
* [Other tools documentation](.) - See respective folders

## ⚙️ Environment Requirements

### General Requirements
- **Operating System**: Windows (most tools based on COM interface)
- **Python**: 3.7 or higher
- **Microsoft Office**: Word and Excel (for COM operations)

### Python Dependencies
```bash
pip install pandas pywin32 openpyxl python-docx win32com
```

### Initial Setup
```bash
# Complete pywin32 configuration (first time only)
python -m pip install --upgrade pywin32
python Scripts/pywin32_postinstall.py -install
```

## ⚠️ Important Notes

1. **Backup Files**: Always backup original Word and Excel files before running any automation
2. **Close Files**: Close all related Word and Excel files before running to avoid permission conflicts
3. **Verify Configuration**: Before execution, verify Excel configuration field names and paths
4. **First Run**: First run may be slow (COM initialization), please be patient

## 🔧 Technology Stack

| Technology | Purpose |
|------------|---------|
| Python 3.x | Main development language |
| pandas | Data processing |
| python-docx | Word file operations |
| pywin32 | Windows COM interface |
| openpyxl | Excel operations |

## 📝 Version History & Cleanup

### Convergence & Field Spec (2026-10)
- ✅ `report_generator` is now the single recommended report generator
- ✅ `legacy/integrated_report`, `legacy/piero_v6`, `legacy/v33_claude` marked **legacy** (reference only)
- ✅ Added `Master_Log.xlsx` field specification (see above)

### Privacy Sanitization (2026-10)
- ✅ Removed internal report templates and sample data files (contained internal project information)
- ✅ Input files (`Master_Log.xlsx`, `Template.doc`) must be provided by the user

### Latest Cleanup (2024-01)
- ✅ Deleted obsolete files and duplicate versions
- ✅ Renamed files for standardization
- ✅ Added detailed README for all tools
- ✅ Unified version marking
- ✅ Updated root documentation
- ✅ Organized duplicate code

**Cleanup Results**:
- Streamlined from 15+ folders to 10 effective folders
- Deleted 27 obsolete files
- Deleted 12 old version outputs
- Added detailed documentation for 9 folders

## 🤝 Contributors

- **Luke**: Format-preserving replacement logic, outline filter tool optimization
- **Piero**: Core report generation, date filling tool, table management
- **Claude AI**: New version experiments, integration optimization, documentation

## 📞 Technical Support

For issues or suggestions:
1. Check the README.md documentation for each tool
2. Verify Excel configuration file format
3. Verify Word and Excel files are properly closed
4. Contact the respective tool developer

## 📄 License & Usage

This project content is for internal use. Please ensure you understand the requirements and limitations of each tool before use.

---

**Recommended Workflow**: 
1. Use `report_generator` (the single recommended entry point)
2. Use standalone utilities (date filling, outline filter) as needed
3. Legacy generators (`legacy/integrated_report`, `legacy/piero_v6`) are kept for reference only
