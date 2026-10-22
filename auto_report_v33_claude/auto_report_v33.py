import pandas as pd
import os
import win32com.client as win32
import re

# --- 路徑設定 ---
DATA_DIR = r'.'
SCRIPT_DIR = r'.'
EXCEL_NAME = 'Master_Log.xlsx (2).xlsx'
TEMPLATE_NAME = 'Template (1).doc' 

excel_path = os.path.abspath(os.path.join(DATA_DIR, EXCEL_NAME))
template_path = os.path.abspath(os.path.join(DATA_DIR, TEMPLATE_NAME))
final_output = os.path.abspath(os.path.join(SCRIPT_DIR, 'Final_Report.docx'))

def process_report_v33_smart():
    if not os.path.exists(excel_path) or not os.path.exists(template_path):
        print("錯誤：找不到檔案。")
        return

    # --- 1. 讀取 Excel 數據 ---
    print("正在讀取 Excel 數據...")
    
    # 🎯 智慧功能 A：從第一個工作表的 A 欄精準抓取報告基本資訊 (免改 Excel)
    try:
        df_all = pd.read_excel(excel_path, sheet_name=0, header=None)
        info_data = {
            'Report_Title': str(df_all.iloc[1, 0]).strip() if pd.notna(df_all.iloc[1, 0]) else "",
            'Doc_Number':   str(df_all.iloc[2, 0]).strip() if pd.notna(df_all.iloc[2, 0]) else "",
            'Doc_Version':  str(df_all.iloc[3, 0]).strip() if pd.notna(df_all.iloc[3, 0]) else "",
            'Model_Name':    str(df_all.iloc[4, 0]).strip() if pd.notna(df_all.iloc[4, 0]) else "",
            'Part_Number':  str(df_all.iloc[5, 0]).strip() if pd.notna(df_all.iloc[5, 0]) else "",
            'Report_Date':  str(df_all.iloc[6, 0]).strip() if pd.notna(df_all.iloc[6, 0]) else "",
            'Author':       str(df_all.iloc[7, 0]).strip() if pd.notna(df_all.iloc[7, 0]) else ""
        }
        print("💡 [智慧識別成功] 已成功從 Excel 抓取基本資訊：")
        for k, v in info_data.items():
            print(f"  - {k}: {v}")
    except Exception as e:
        print(f"❌ 剖析 Excel 基本資訊失敗，錯誤原因: {e}")
        return

    # 讀取主要測項 Log
    df = pd.read_excel(excel_path, sheet_name=0)
    # 🎯 清理 Excel 欄位名稱的頭尾空白，避免因空白導致找不到欄位
    if hasattr(df.columns, 'str'):
        df.columns = df.columns.str.strip()
    
    # 建立跳過清單
    skipped_items = df[df['測試狀態'].str.contains('跳過', na=False)]['測項名稱'].dropna().unique().tolist()
    
    # 建立時間與結果映射表 (核心測項名稱 -> 時間與結果資料)
    time_map = {}
    result_map = {}
    for _, row in df.iterrows():
        full_name = str(row['測項名稱']).strip()
        core_name = re.sub(r'^[0-9. \t]+', '', full_name).strip()
        
        start = pd.to_datetime(row['開始時間']).strftime('%Y/%m/%d') if pd.notna(row['開始時間']) else ""
        end = pd.to_datetime(row['結束時間']).strftime('%Y/%m/%d') if pd.notna(row['結束時間']) else ""
        
        time_map[core_name] = f"{start}~{end}" if start or end else ""
        
        pd_val = str(row.get('Physical Damage', '')).strip() if pd.notna(row.get('Physical Damage', '')) else ""
        fc_val = str(row.get('Functional Check', '')).strip() if pd.notna(row.get('Functional Check', '')) else ""
        
        if pd_val.lower() == 'fail' or fc_val.lower() == 'fail':
            final_result = "Fail"
        elif pd_val.lower() == 'pass' or fc_val.lower() == 'pass':
            final_result = "Pass"
        else:
            final_result = ""
            
        result_map[core_name] = {
            'Physical_Damage': pd_val,
            'Functional_Check': fc_val,
            'Result': final_result
        }

    # --- 2. 開啟 Word ---
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(template_path)
    toc_end = doc.TablesOfContents(1).Range.End if doc.TablesOfContents.Count > 0 else 0

    # --- 🌟 智慧功能 B：全方位穿透置換變數（包含表頭頁首頁尾 + 消除黃色高亮） 🌟 ---
    print("\n[智慧欄位置換] 開始進行內文與表頭的欄位置換...")
    replacements = {
        "{Report_Title}": info_data['Report_Title'],
        "{Doc_Number}":   info_data['Doc_Number'],
        "{Doc_Version}":  info_data['Doc_Version'],
        "{Model_Name}":   info_data['Model_Name'],
        "{Part_Number}":  info_data['Part_Number'],
        "{Report_Date}":  info_data['Report_Date'],
        "{Author}":       info_data['Author']
    }

    # 精準置換與除高亮的副程式
    def execute_replace_on_range(rng, find_str, replace_str):
        find_obj = rng.Find
        find_obj.ClearFormatting()
        find_obj.Text = find_str
        find_obj.Replacement.ClearFormatting()
        find_obj.Replacement.Highlight = False  # 💡 自動把原本的黃色高亮清除
        find_obj.Replacement.Text = replace_str
        find_obj.Execute(Replace=2) # 2 = wdReplaceAll

    for target_tag, real_value in replacements.items():
        # 1. 替換主要內文層級
        for story in doc.StoryRanges:
            execute_replace_on_range(story, target_tag, real_value)
            
        # 2. 🎯 穿透 Section 結構，強制替換表頭 (Header) 與頁尾 (Footer) 裡面的標籤
        for section in doc.Sections:
            for header in section.Headers:
                if header.Exists:
                    execute_replace_on_range(header.Range, target_tag, real_value)
            for footer in section.Footers:
                if footer.Exists:
                    execute_replace_on_range(footer.Range, target_tag, real_value)
        print(f"  - 標籤 {target_tag} ➔ 穿透置換完成")

    # --- 第一階段：處理移除 (原 V30 穩定邏輯) ---
    print(f"\n開始執行移除邏輯...")
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

    # --- 第三階段：精準填入 3.1 表格時間與結果 ---
    print("\n開始填入 3.1 Test Item Result 時間與結果...")
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
                    
                    if "test item" in col2_header:
                        for r in range(2, tbl.Rows.Count + 1):
                            try:
                                cell_text = tbl.Cell(r, 2).Range.Text
                                clean_cell_name = re.sub(r'[\r\x07\t]', '', cell_text).strip()
                                
                                if not clean_cell_name: continue

                                if clean_cell_name in time_map:
                                    time_str = time_map[clean_cell_name]
                                    if time_str:
                                        tbl.Cell(r, 3).Range.Text = time_str
                                        print(f"  - [精確更新時間] {clean_cell_name} -> {time_str}")
                                
                                if clean_cell_name in result_map:
                                    res_val = result_map[clean_cell_name]['Result']
                                    if res_val:
                                        tbl.Cell(r, 4).Range.Text = res_val
                                        print(f"  - [精確更新結果] {clean_cell_name} -> {res_val}")
                            except: continue
                        break 
                except: continue

    # --- 第 3.5 階段：精準填寫第 4 章節各測項詳細 Result 表格 ---
    print("\n開始填寫第 4 章節各測項詳細 Result...")
    for tbl in doc.Tables:
        try:
            h1 = tbl.Cell(1, 1).Range.Text.strip().lower()
            h2 = tbl.Cell(1, 2).Range.Text.strip().lower()
            if "test result" in h1 and "before test" in h2:
                # 這是我們要找的 Result 表格。往上找最近的 Level 3 Heading 確定是哪個測項
                rng = tbl.Range
                rng.Collapse(Direction=1)
                
                test_name = ""
                for _ in range(50):
                    rng.Move(Unit=4, Count=-1) # Move backward 1 paragraph
                    if rng.Paragraphs.Count > 0:
                        p = rng.Paragraphs(1)
                        if p.OutlineLevel == 3: # 第 4 章的測項標題皆為 Level 3
                            raw_heading = p.Range.Text.strip().replace('\r', '').replace('\x07', '')
                            test_name = re.sub(r'^[0-9. \t]+', '', raw_heading).strip()
                            break
                
                if test_name and test_name in result_map:
                    res_data = result_map[test_name]
                    for r in range(2, tbl.Rows.Count + 1):
                        row_header = tbl.Cell(r, 1).Range.Text.strip().replace('\r', '').replace('\x07', '')
                        
                        if "physical damage" in row_header.lower() and res_data['Physical_Damage']:
                            tbl.Cell(r, 2).Range.Text = "Pass"
                            tbl.Cell(r, 3).Range.Text = res_data['Physical_Damage']
                            print(f"  - [{test_name}] Physical Damage -> Before: Pass, After: {res_data['Physical_Damage']}")
                            
                        if "functional check" in row_header.lower() and res_data['Functional_Check']:
                            tbl.Cell(r, 2).Range.Text = "Pass"
                            tbl.Cell(r, 3).Range.Text = res_data['Functional_Check']
                            print(f"  - [{test_name}] Functional Check -> Before: Pass, After: {res_data['Functional_Check']}")
        except:
            pass

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

    # 💡 智慧功能 C：全文字公式、頁碼與目錄刷新
    print("\n正在強制更新目錄與全文字公式...")
    word.ActiveWindow.View.Type = 3  
    doc.Fields.Update()
    for story in doc.StoryRanges:
        story.Fields.Update()
    doc.Fields.Update()

    doc.SaveAs(final_output, FileFormat=16)
    print(f"\n【完美達成】最終檔案已生成於：{final_output}")

if __name__ == "__main__":
    process_report_v33_smart()