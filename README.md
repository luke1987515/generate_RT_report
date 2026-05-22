# Generate RT Report - Complete Word Automation Solution

[English](README.md) | [繁體中文](README.zh-TW.md)

A complete automated Microsoft Word RT report generation solution. Includes report generation, table management, outline processing and more, with support for batch and intelligent report generation workflows.

## 🎯 Core Features

### Production Tools (Recommended)

| Tool | Description | Version | Status |
|------|-------------|---------|--------|
| **integrated_report_tool** | ⭐ Latest integrated version with all features | v1.0 | ✅ Recommended |
| **generate_RT_report_piero** | Core report generation (v6 stable) | v6 | ✅ Stable |
| **auto_report-table-date_piero** | Table date filling & section management | v1.0 | ✅ Stable |
| **word_outline_filter** | Intelligent outline filtering | v8.0 | ✅ Latest |

### Supplementary Tools

| Tool | Description | Purpose |
|------|-------------|---------|
| **generate_RT_report_luke** | Format-preserving replacement logic | Run-level fine control |
| **auto_report_v33_claude** | AI-assisted report generation (experimental) | Intelligent information recognition |
| **auto_report-table_modify_piero** | Table content modification tool | Modify table values and status |
| **find_tables_in_word** | Table detection and analysis | Table location and structure analysis |
| **read_word_outline** | Outline reading tool | Document structure parsing |
| **doc_to_template** | Document & template conversion | File format conversion |

## 📁 Project Structure

```
generate_RT_report/
├── integrated_report_tool/           # ⭐ Recommended first choice
│   ├── unified_report.py             # Integrated main program
│   ├── create_config.py              # Config creation tool
│   └── Report_Config.xlsx            # Multi-sheet configuration
│
├── generate_RT_report_piero/         # v6 stable version
│   ├── auto_report.py
│   ├── create_config.py
│   └── Master_Log.xlsx
│
├── generate_RT_report_luke/          # Format preservation reference
│   ├── docx_replace.py               # Run-level replacement
│   └── requirements.txt
│
├── auto_report-table-date_piero/     # Date filling stable version
│   ├── auto_report_table_date.py
│   └── Master_Log.xlsx
│
├── auto_report-table_modify_piero/   # Table modification tool
│   └── auto_report_table_modify.py
│
├── auto_report_v33_claude/           # AI experimental version
│   └── auto_report_v33.py
│
├── word_outline_filter/              # v8 latest version
│   └── word_outline_filter.py
│
├── read_word_outline/                # Outline reading
│   └── read_word_outline.py
│
├── find_tables_in_word/              # Table detection
│   └── find_tables_in_word.py
│
└── doc_to_template/                  # Format conversion
    ├── doc_to_template.py
    └── template_to_doc.py
```

## 🚀 Quick Start

### Option A: Using Integrated Tool (Recommended)

```bash
cd integrated_report_tool
python create_config.py           # Create configuration
python unified_report.py          # Generate report
```

### Option B: Using Piero Stable Version

```bash
cd generate_RT_report_piero
python auto_report.py             # Generate report
```

### Option C: Using Date Filling Tool

```bash
cd auto_report-table-date_piero
python auto_report_table_date.py  # Fill dates and times
```

## 📋 Version Comparison & Selection Guide

### Which version to use?

| Use Case | Recommended Tool | Reason |
|----------|------------------|--------|
| New project, need complete features | **integrated_report_tool** | Most complete, latest, best maintainability |
| Need simple report generation | generate_RT_report_piero | Stable, reliable, feature-complete |
| Only need date/time filling | auto_report-table-date_piero | Dedicated tool, precise functionality |
| Need fine format control | generate_RT_report_luke | Run-level replacement, complete format preservation |
| Need to experiment with AI features | auto_report_v33_claude | Latest AI-assisted, feature-rich |
| Need to analyze table structure | find_tables_in_word | Dedicated table analysis tool |
| Need to filter outline content | word_outline_filter | Latest version (v8.0), powerful features |

## 🔄 Workflow Example

### Complete Automation Workflow

```
1. Prepare Excel configuration (Master_Log.xlsx or Report_Config.xlsx)
   ↓
2. Prepare Word template (Template.doc or Template.docx)
   ↓
3. Run report generation
   - Use integrated_report_tool (recommended)
   - Or generate_RT_report_piero
   ↓
4. (Optional) Use auto_report-table-date_piero to fill dates
   ↓
5. (Optional) Use word_outline_filter to filter content
   ↓
6. Final report: Final_Report.docx
```

## 📚 Detailed Documentation

Detailed documentation for each tool can be found in the README.md files in their respective folders:

* [integrated_report_tool/README.md](integrated_report_tool/README.md) - Integrated tool complete documentation
* [generate_RT_report_piero/README.md](generate_RT_report_piero/README.md) - Piero version documentation
* [generate_RT_report_luke/README.md](generate_RT_report_luke/README.md) - Luke version documentation
* [auto_report-table-date_piero/README.md](auto_report-table-date_piero/README.md) - Date tool documentation
* [word_outline_filter/README.md](word_outline_filter/README.md) - Outline filter tool documentation
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
1. First choice: `integrated_report_tool` (latest and most complete)
2. Alternative: `generate_RT_report_piero` (stable and reliable)
3. Combine other tools based on specific needs
