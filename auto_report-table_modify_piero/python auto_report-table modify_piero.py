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

def remove_skipped_sections_v30():
    if not os.path.exists(excel_path) or not os.path.exists(template_path):
        print("錯誤：找不到檔案。")
        return

    df = pd.read_excel(excel_path)
    skipped_items = df[df['測試狀態'].str.contains('跳過', na=False)]['測項名稱'].dropna().unique().tolist()
    print(f"預計移除項目: {skipped_items}")

    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True
    doc = word.Documents.Open(template_path)
    toc_end = doc.TablesOfContents(1).Range.End if doc.TablesOfContents.Count > 0 else 0

    # --- 第一階段：處理表格內合併區塊刪除 ---
    for full_name in skipped_items:
        core_name = re.sub(r'^[0-9. \t]+', '', full_name).strip()
        search_key = core_name[:15] 
        print(f"\n[處理中] {core_name}")

        find_rng = doc.Range(toc_end, doc.Content.End)
        while find_rng.Find.Execute(FindText=search_key):
            curr_p = find_rng.Paragraphs(1)
            
            if curr_p.Range.Information(win32.constants.wdWithInTable):
                try:
                    # 💡 關鍵修正：完全避開 Rows(i)
                    # 1. 取得目標儲存格並選取它
                    target_cell = curr_p.Range.Cells(1)
                    target_cell.Select()
                    
                    # 2. 使用 Selection.Rows.Delete() 
                    # 這是 Word 的「黑科技」：如果選取的格子是垂直合併的，
                    # 呼叫 Selection.Rows.Delete 會連同該合併格對應的所有行一起刪除。
                    word.Selection.Rows.Delete()
                    print(f"   -> [表格] 已移除合併區塊（包含右側所有關聯行）")
                except Exception as e:
                    print(f"   -> [警告] 刪除失敗，嘗試清空內容...")
                    curr_p.Range.Text = ""
                
                # 刪除後表格結構改變，從頭搜尋確保安全
                find_rng = doc.Range(toc_end, doc.Content.End)
                continue
            
            # --- 章節標題刪除 (保留 V12 原則) ---
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

    # --- 第二階段：深度掃描空標題 (V12 原則) ---
    print("\n[清理] 執行深度空標題掃描...")
    for _ in range(3):
        deleted_count = 0
        for i in range(doc.Paragraphs.Count, 0, -1):
            try:
                p = doc.Paragraphs(i)
                if p.Range.Start < toc_end: continue
                if p.OutlineLevel < 10:
                    is_empty = True
                    scan_p = p.Next()
                    while scan_p:
                        if scan_p.OutlineLevel <= p.OutlineLevel: break
                        clean_text = scan_p.Range.Text.strip().replace('\r', '').replace('\x0c', '').replace('\t', '')
                        if len(clean_text) > 0 or scan_p.Range.Information(win32.constants.wdWithInTable):
                            is_empty = False
                            break
                        scan_p = scan_p.Next()
                    if is_empty:
                        t_text = p.Range.Text.strip()
                        if len(t_text) > 0 and not any(x in t_text for x in ["Revision History", "UUT Configuration", "Appendix"]):
                            p.Range.Delete()
                            deleted_count += 1
            except: continue
        if deleted_count == 0: break

    # --- 第三階段：序號重編 (安全模式) ---
    print("\n[序號] 重新編號...")
    for table in doc.Tables:
        try:
            # 避開 table.Rows，改用 table.Range.Cells 遍歷
            first_cell_text = table.Range.Cells(1).Range.Text.lower()
            if any(k in first_cell_text for k in ["no", "序號"]):
                count = 1
                # 重新計算總格數，避免結構變動導致報錯
                total_cells = table.Range.Cells.Count
                for c_idx in range(1, total_cells + 1):
                    try:
                        c = table.Range.Cells(c_idx)
                        if c.ColumnIndex == 1 and c.RowIndex > 1:
                            c.Range.Text = str(count)
                            count += 1
                    except: continue
        except: continue

    doc.Fields.Update()
    doc.SaveAs(final_output, FileFormat=16)
    print(f"任務成功！檔案儲存至: {final_output}")

if __name__ == "__main__":
    remove_skipped_sections_v30()