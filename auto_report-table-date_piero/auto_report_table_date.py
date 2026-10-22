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

def process_report_v33():
    if not os.path.exists(excel_path) or not os.path.exists(template_path):
        print("錯誤：找不到檔案。")
        return

    # --- 1. 讀取 Excel 數據 ---
    df = pd.read_excel(excel_path)
    
    # 建立跳過清單
    skipped_items = df[df['測試狀態'].str.contains('跳過', na=False)]['測項名稱'].dropna().unique().tolist()
    
    # 建立時間映射表 (核心測項名稱 -> 時間字串)
    time_map = {}
    for _, row in df.iterrows():
        # 移除序號後的純名稱作為 key，用於精確比對
        full_name = str(row['測項名稱']).strip()
        core_name = re.sub(r'^[0-9. \t]+', '', full_name).strip()
        
        start = pd.to_datetime(row['開始時間']).strftime('%Y/%m/%d') if pd.notna(row['開始時間']) else ""
        end = pd.to_datetime(row['結束時間']).strftime('%Y/%m/%d') if pd.notna(row['結束時間']) else ""
        
        time_map[core_name] = f"{start}~{end}" if start or end else ""

    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True
    doc = word.Documents.Open(template_path)
    toc_end = doc.TablesOfContents(1).Range.End if doc.TablesOfContents.Count > 0 else 0

    # --- 第一階段：處理移除 (原 V30 穩定邏輯) ---
    print(f"開始執行移除邏輯...")
    for full_name in skipped_items:
        core_name = re.sub(r'^[0-9. \t]+', '', full_name).strip()
        search_key = core_name[:15] 

        find_rng = doc.Range(toc_end, doc.Content.End)
        while find_rng.Find.Execute(FindText=search_key):
            curr_p = find_rng.Paragraphs(1)
            if curr_p.Range.Information(win32.constants.wdWithInTable):
                try:
                    target_cell = curr_p.Range.Cells(1)
                    target_cell.Select()
                    word.Selection.Rows.Delete()
                    print(f"  - [表格項移除] {core_name}")
                except:
                    curr_p.Range.Text = ""
                find_rng = doc.Range(toc_end, doc.Content.End)
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
                print(f"  - [章節移除] {core_name}")
                find_rng = doc.Range(start_pos, doc.Content.End)
            else:
                find_rng = doc.Range(find_rng.End, doc.Content.End)

    # --- 第二階段：清理空標題 ---
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

    # --- 第三階段：精準填入 3.1 表格時間 (嚴格匹配版) ---
    print("\n開始填入 3.1 Test Item Result 時間 (嚴格比對)...")
    
    start_31 = 0
    for p in doc.Paragraphs:
        if "3.1" in p.Range.Text and "Test Item Result" in p.Range.Text:
            start_31 = p.Range.Start
            break
    
    if start_31 > 0:
        for tbl in doc.Tables:
            if tbl.Range.Start >= start_31:
                try:
                    col2_header = tbl.Cell(1, 2).Range.Text.lower()
                    col3_header = tbl.Cell(1, 3).Range.Text.lower()
                    
                    if "test item" in col2_header and "date" in col3_header:
                        for r in range(2, tbl.Rows.Count + 1):
                            try:
                                # 取得並徹底清理 Word 儲存格文字
                                cell_text = tbl.Cell(r, 2).Range.Text
                                clean_cell_name = re.sub(r'[\r\x07\t]', '', cell_text).strip()
                                
                                if not clean_cell_name: continue

                                # 🚩 使用 == 進行嚴格比對，避免 Operating 誤刷到 Non-Operating
                                if clean_cell_name in time_map:
                                    time_str = time_map[clean_cell_name]
                                    tbl.Cell(r, 3).Range.Text = time_str
                                    print(f"  - [精確更新] {clean_cell_name} -> {time_str}")
                            except: continue
                        break 
                except: continue

    # --- 第四階段：重新編號 (安全模式) ---
    print("\n[序號] 重新編號...")
    for table in doc.Tables:
        try:
            first_cell_text = table.Range.Cells(1).Range.Text.lower()
            if any(k in first_cell_text for k in ["no", "序號"]):
                count = 1
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
    print(f"\n任務完成！最終檔案：{final_output}")

if __name__ == "__main__":
    process_report_v33()