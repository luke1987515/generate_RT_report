import pandas as pd
import os
import win32com.client as win32
import re

# --- 路徑設定 ---
DATA_DIR = r'D:\DQA_Templates'
SCRIPT_DIR = r'D:\Python Script\report transfer'
EXCEL_NAME = 'Master_Log.xlsx.xlsx'
TEMPLATE_NAME = 'Template.doc' 

excel_path = os.path.join(DATA_DIR, EXCEL_NAME)
template_path = os.path.join(DATA_DIR, TEMPLATE_NAME)
final_output = os.path.join(SCRIPT_DIR, 'Final_Report.docx')

def remove_skipped_sections_v12():
    if not os.path.exists(excel_path) or not os.path.exists(template_path):
        print("錯誤：找不到檔案。")
        return

    df = pd.read_excel(excel_path)
    skipped_items = df[df['測試狀態'].str.contains('跳過', na=False)]['測項名稱'].dropna().unique().tolist()
    print(f"預計移除項目: {skipped_items}")

    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True
    doc = word.Documents.Open(template_path)

    toc_end = 0
    if doc.TablesOfContents.Count > 0:
        toc_end = doc.TablesOfContents(1).Range.End

    # --- 第一階段：詳解內容區間刪除 ---
    for full_name in skipped_items:
        core_name = re.sub(r'^[0-9. \t]+', '', full_name).strip()
        search_key = core_name[:15] 
        print(f"\n[處理中] {core_name}")

        find_rng = doc.Range(toc_end, doc.Content.End)
        while find_rng.Find.Execute(FindText=search_key):
            curr_p = find_rng.Paragraphs(1)
            if curr_p.Range.Information(win32.constants.wdWithInTable):
                try: curr_p.Range.Rows.Delete()
                except: pass
                find_rng = doc.Range(find_rng.End, doc.Content.End)
                continue
            
            if curr_p.OutlineLevel < 10:
                target_level = curr_p.OutlineLevel
                start_pos = curr_p.Range.Start
                next_p = curr_p.Next()
                end_pos = doc.Content.End
                while next_p:
                    if next_p.OutlineLevel <= target_level:
                        if search_key[:10].lower() not in next_p.Range.Text.lower():
                            end_pos = next_p.Range.Start
                            break
                    next_p = next_p.Next()
                doc.Range(start_pos, end_pos).Delete()
                find_rng = doc.Range(start_pos, doc.Content.End)
            else:
                find_rng = doc.Range(find_rng.End, doc.Content.End)

    # --- 第二階段：🚩 深度掃描並刪除「空的大項」 (Other 等) ---
    print("\n[清理] 執行深度空標題掃描...")
    
    # 執行多次掃描以處理連鎖空項 (例如刪掉 4.1.2 後 4.1 也變空的情況)
    for _ in range(3):
        deleted_count = 0
        for i in range(doc.Paragraphs.Count, 0, -1):
            try:
                p = doc.Paragraphs(i)
                if p.Range.Start < toc_end: continue # 避開目錄
                
                if p.OutlineLevel < 10: # 如果是標題 (1-9級)
                    is_empty = True
                    scan_p = p.Next()
                    
                    # 往下掃描，直到遇到下一個「同級或更高級」標題
                    while scan_p:
                        # 如果遇到同級或上級標題，代表中間沒內容，確定為空
                        if scan_p.OutlineLevel <= p.OutlineLevel:
                            is_empty = True
                            break
                        
                        # 檢查內容：移除掉換行、空格、分頁符號後，還有字嗎？
                        clean_text = scan_p.Range.Text.strip().replace('\r', '').replace('\x0c', '').replace('\t', '')
                        
                        # 檢查是否有表格
                        in_table = scan_p.Range.Information(win32.constants.wdWithInTable)
                        
                        if len(clean_text) > 0 or in_table:
                            # 發現了實質內容（或者是子標題），不能刪除此標題
                            is_empty = False
                            break
                        scan_p = scan_p.Next()
                    
                    if is_empty:
                        t_text = p.Range.Text.strip()
                        # 排除關鍵保留區
                        if len(t_text) > 0 and not any(x in t_text for x in ["Revision History", "UUT Configuration", "Appendix"]):
                            print(f"   -> [自動移除空項] {t_text[:30]}")
                            p.Range.Delete()
                            deleted_count += 1
            except:
                continue
        
        if deleted_count == 0:
            break

    print("\n[完成] 更新並儲存...")
    doc.Fields.Update()
    doc.SaveAs(final_output, FileFormat=16)
    print(f"任務成功！")

if __name__ == "__main__":
    remove_skipped_sections_v12()