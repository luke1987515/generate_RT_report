# Generate RT Report

[English](README.md) | [繁體中文](README.zh-TW.md)

這個儲存庫包含了一系列用於自動化產生與處理 Microsoft Word (RT) 報告的腳本與工具。

## 功能特色

本專案分為幾個子模組，分別針對報告自動化的不同方面：

* **自動化 Word 報告生成**：用於產生具備特定格式與結構報告的腳本 (`generate_RT_report_*`)。
* **表格修改與時間戳記**：用於修改 Word 文件內表格，以及記錄/更新時間的工具 (`auto_report-table-date_piero`, `auto_report-table_modify_piero`)。
* **大綱處理**：用於讀取與過濾 Word 文件大綱的公用程式 (`read_word_outline`, `word_outline_filter`)。
* **範本管理**：將文件轉換為範本 (`doc_to_template`)。
* **整合工具**：結合多項功能的綜合性報告工具 (`integrated_report_tool`)。

## 專案結構

* `auto_report-table-date_piero/`：用於管理報告表格中日期與時間的腳本。
* `auto_report-table_modify_piero/`：用於一般表格修改的腳本。
* `doc_to_template/`：將文件轉為範本的公用程式。
* `find_tables_in_word/`：用於在 Word 文件中尋找並萃取表格的工具。
* `generate_RT_report_luke/` & `generate_RT_report_piero/`：由不同貢獻者開發的核心報告生成腳本。
* `integrated_report_tool/`：結合多項功能的綜合工具。
* `read_word_outline/` & `word_outline_filter/`：用於處理 Word 大綱與章節的腳本。

## 開始使用

每個目錄都包含特定的工具與腳本。請進入個別的目錄以獲取更詳細的說明或與該任務相關的腳本。
