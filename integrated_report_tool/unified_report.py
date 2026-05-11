import pandas as pd
import os
import win32com.client as win32
import re

# --- 路徑設定 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, 'Report_Config.xlsx')
# 這裡預設範本放在同資料夾，名稱為 Template.docx (請使用者將範本放入)
TEMPLATE_PATH = os.path.join(BASE_DIR, 'Template.docx') 
OUTPUT_PATH = os.path.join(BASE_DIR, 'Final_Report_Unified.docx')

def unified_report_process():
    global TEMPLATE_PATH
    if not os.path.exists(CONFIG_PATH):
        print(f"錯誤：找不到設定檔 {CONFIG_PATH}")
        return
    
    if not os.path.exists(TEMPLATE_PATH):
        # 如果找不到 Template.docx，嘗試找 Template.doc (Piero 的版本)
        alt_template = os.path.join(BASE_DIR, 'Template.doc')
        if os.path.exists(alt_template):
            TEMPLATE_PATH = alt_template
            print(f"提示：找不到預設的 Template.docx，自動切換至 {TEMPLATE_PATH}")
        else:
            print(f"錯誤：請確保範本檔案 (Template.docx 或 Template.doc) 已放置於 {BASE_DIR}")
            return

    print("正在讀取 Excel 設定...")
    df_replacements = pd.read_excel(CONFIG_PATH, sheet_name='Replacements')
    df_status = pd.read_excel(CONFIG_PATH, sheet_name='Test_Status')
    
    # 建立替換字典
    replacements = dict(zip(df_replacements.iloc[:, 0].astype(str), df_replacements.iloc[:, 1].astype(str)))
    
    # 獲取需要刪除的項目 (狀態包含 '跳過' 或 'Skipped')
    skipped_items = df_status[df_status['測試狀態'].str.contains('跳過|Skipped', na=False, case=False)]['測項名稱'].dropna().unique().tolist()

    print(f"預計替換項目數量: {len(replacements)}")
    print(f"預計刪除章節: {skipped_items}")

    # 啟動 Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True
    doc = None
    
    try:
        doc = word.Documents.Open(TEMPLATE_PATH)

        # --- 階段 1: 文字替換 (Luke 的邏輯) ---
        print("\n[階段 1] 執行文字替換...")
        for tag, value in replacements.items():
            print(f"   -> 替換 {tag} 為 {value}")
            # 使用 Word 內建的尋找與替換 (全域)
            find_range = doc.Content
            find_range.Find.ClearFormatting()
            find_range.Find.Replacement.ClearFormatting()
            find_range.Find.Text = tag
            find_range.Find.Replacement.Text = value
            find_range.Find.Execute(Replace=2) # 2 = wdReplaceAll

        # --- 階段 2: 刪除跳過章節 (Piero 的邏輯) ---
        print("\n[階段 2] 執行章節清理...")
        
        # 偵測目錄位置，避開目錄區
        toc_end = 0
        if doc.TablesOfContents.Count > 0:
            toc_end = doc.TablesOfContents(1).Range.End
            print(f"   偵測到目錄，將從位置 {toc_end} 之後開始處理。")

        for full_name in skipped_items:
            # 去除編號，只取核心名稱
            core_name = re.sub(r'^[0-9. \t]+', '', full_name).strip()
            search_key = core_name[:15] # 縮短關鍵字增加搜尋成功率
            print(f"   [清理中] {core_name}")

            keep_searching = True
            while keep_searching:
                find_rng = doc.Range(toc_end, doc.Content.End)
                find_rng.Find.ClearFormatting()
                find_rng.Find.Text = search_key
                
                if find_rng.Find.Execute():
                    curr_p = find_rng.Paragraphs(1)
                    
                    # A: 如果在表格內 (通常是摘要表)
                    if curr_p.Range.Information(win32.constants.wdWithInTable):
                        try:
                            curr_p.Range.Rows.Delete()
                            print(f"      -> 已移除摘要表行")
                        except:
                            curr_p.Range.Select()
                            word.Selection.Rows.Delete()
                        continue
                    
                    # B: 如果是章節標題 (OutlineLevel < 10)
                    if curr_p.OutlineLevel < 10:
                        print(f"      -> 鎖定章節標題，開始刪除內容...")
                        start_anchor = curr_p.Range.Start
                        
                        # 循環刪除段落，直到遇到下一個標題
                        steps = 0
                        while steps < 300: # 避免無限循環
                            steps += 1
                            try:
                                target_p = doc.Range(start_anchor, doc.Content.End).Paragraphs(1)
                                t_text = target_p.Range.Text.strip()
                                t_level = target_p.OutlineLevel
                                
                                # 停止條件：遇到下一個正式標題 且 不是當前項目
                                if steps > 1:
                                    if t_level < 10 and len(t_text) > 0 and search_key[:8] not in t_text:
                                        break
                                    # 碰到特殊區域停止
                                    if any(x in t_text for x in ["Revision History", "UUT Configuration"]):
                                        break
                                
                                target_p.Range.Delete()
                            except:
                                break
                        print(f"      -> 章節內容清理完成")
                    else:
                        # 只是內文引用，跳過
                        toc_end = find_rng.End
                else:
                    keep_searching = False

        # --- 階段 3: 更新與儲存 ---
        print("\n[階段 3] 更新目錄並儲存...")
        doc.Fields.Update() # 更新所有變數與目錄
        doc.SaveAs(OUTPUT_PATH, FileFormat=16) # 16 = wdFormatXMLDocument (.docx)
        print(f"\n任務成功！檔案已儲存至：\n{OUTPUT_PATH}")
    
    finally:
        if doc:
            doc.Close(False) # 不儲存對範本的變更
        word.Quit()

if __name__ == "__main__":
    try:
        unified_report_process()
    except Exception as e:
        print(f"執行出錯: {e}")
