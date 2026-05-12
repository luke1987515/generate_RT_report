# RT 報告自動化工具 (日期填寫版) - auto_report-table-date_piero

這是一個用於自動化產生 Word 測試報告的 Python 工具。它能根據 Excel (Master Log) 的內容，自動處理 Word 範本中的章節刪除、表格列更新以及測試日期填寫。

## 主要功能 (Key Features)

1.  **自動刪除跳過項 (Skip Items)**：
    *   讀取 Excel 中的「測試狀態」，若包含「跳過」字眼，自動刪除 Word 範本中對應的章節內容。
    *   同時會刪除 3.1 表格中對應的測試項目列。
2.  **清理空標題**：
    *   在刪除內容後，自動掃描並移除沒有實質內容的空白標題，保持報告整潔。
3.  **精準填入測試時間 (Section 3.1)**：
    *   針對 Word 報告中的「3.1 Test Item Result」表格。
    *   根據 Excel 中的「開始時間」與「結束時間」，自動填入對應測試項目的「Date」欄位。
    *   使用「嚴格比對」機制，避免名稱相似（如 Operating 與 Non-Operating）造成的誤填。
4.  **自動重新編號 (Table Renumbering)**：
    *   自動更新表格中的「No」或「序號」欄位，確保刪除項目後序號仍保持連續。
5.  **更新目錄**：
    *   自動執行 Word 的欄位更新 (Fields Update)，確保目錄 (TOC) 正確。

## 檔案結構

*   `auto_report-table-date.py`：主要的自動化腳本 (原始檔名可能為 `python auto_report-table-date.py`)。
*   `Master_Log.xlsx (1) 1.xlsx`：來源數據 Excel 檔案。
*   `Template.doc`：Word 報告範本 (需手動放入此資料夾)。
*   `Final_Report.docx`：執行後產出的最終報告檔案。

## 環境需求

*   **作業系統**：Windows (需安裝 Microsoft Word)。
*   **Python 版本**：Python 3.x。
*   **必要套件**：
    ```bash
    pip install pandas pywin32 openpyxl
    ```

## 使用說明

1.  **準備檔案**：
    *   將 `Template.doc` 放到此資料夾中。
    *   確保 `Master_Log` 的欄位名稱包含「測項名稱」、「測試狀態」、「開始時間」、「結束時間」。
2.  **修改路徑 (重要)**：
    目前的腳本中使用了絕對路徑，建議在執行前修改腳本開頭的設定：
    ```python
    # 建議修改為以下方式以支援目前資料夾
    DATA_DIR = os.getcwd()
    SCRIPT_DIR = os.getcwd()
    EXCEL_NAME = 'Master_Log.xlsx (1) 1.xlsx' # 或是改回 Master_Log.xlsx
    TEMPLATE_NAME = 'Template.doc'
    ```
3.  **執行腳本**：
    在終端機執行：
    ```bash
    python "auto_report-table-date.py"
    ```

## 注意事項

*   執行時請關閉相關的 Excel 與 Word 檔案，避免權限衝突。
*   腳本會自動開啟 Word 視窗進行操作，請勿在執行期間手動操作 Word。
