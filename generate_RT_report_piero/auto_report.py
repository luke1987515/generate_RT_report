import pandas as pd
import os
import win32com.client as win32
import re

# ## --- 路徑設定 ---
# # DATA_DIR = r'D:\DQA_Templates'
# # SCRIPT_DIR = r'D:\Python Script\report transfer'
# # EXCEL_NAME = 'Master_Log.xlsx.xlsx'
# # TEMPLATE_NAME = 'Template.doc' 

# # excel_path = os.path.join(DATA_DIR, EXCEL_NAME)
# # template_path = os.path.join(DATA_DIR, TEMPLATE_NAME)
# # final_output = os.path.join(SCRIPT_DIR, 'Final_Report.docx')

# --- 動態路徑設定 ---
# 獲取目前 Python 腳本所在的資料夾路徑
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 假設你的 Excel 和 Template 都在跟腳本同一個資料夾下
DATA_DIR = BASE_DIR 
# 假設輸出的檔案也要放在同一個資料夾下
SCRIPT_DIR = BASE_DIR 

EXCEL_NAME = 'Master_Log.xlsx'
TEMPLATE_NAME = 'Template.doc' 

# 組合路徑
excel_path = os.path.join(DATA_DIR, EXCEL_NAME)
template_path = os.path.join(DATA_DIR, TEMPLATE_NAME)
final_output = os.path.join(SCRIPT_DIR, 'Final_Report.docx')



def remove_skipped_sections_v6():
    if not os.path.exists(excel_path) or not os.path.exists(template_path):
        print("錯誤：找不到檔案。")
        return

    df = pd.read_excel(excel_path)
    skipped_items = df[df['測試狀態'].str.contains('跳過', na=False)]['測項名稱'].dropna().unique().tolist()
    print(f"預計移除項目: {skipped_items}")

    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True
    doc = word.Documents.Open(template_path)

    # 獲取目錄結束位置，避開保護區
    toc_end = 0
    if doc.TablesOfContents.Count > 0:
        toc_end = doc.TablesOfContents(1).Range.End
        print(f"偵測到目錄，將從位置 {toc_end} 之後開始處理。")

    for full_name in skipped_items:
        core_name = re.sub(r'^[0-9. \t]+', '', full_name).strip()
        search_key = core_name[:15] 
        print(f"\n[處理中] {core_name}")

        keep_searching = True
        while keep_searching:
            # 永遠從目錄後開始找
            find_rng = doc.Range(toc_end, doc.Content.End)
            find_rng.Find.ClearFormatting()
            find_rng.Find.Text = search_key
            
            if find_rng.Find.Execute():
                curr_p = find_rng.Paragraphs(1)
                
                # --- 情況 A: 在表格裡 (摘要表) ---
                if curr_p.Range.Information(win32.constants.wdWithInTable):
                    print(f"   -> [表格] 移除摘要表行")
                    try:
                        # 修正處：必須透過 Range 存取 Rows
                        curr_p.Range.Rows.Delete()
                    except:
                        # 如果 Rows.Delete 失敗，嘗試刪除該段落所在的整格或整列
                        curr_p.Range.Select()
                        word.Selection.Rows.Delete()
                    continue
                
                # --- 情況 B: 內文詳解區 ---
                # 只有大綱層級 < 10 (標題 1-9) 才視為章節起點
                if curr_p.OutlineLevel >= 10:
                    print(f"   -> [跳過] 僅為內文引用")
                    toc_end = find_rng.End # 往後移，避免重複抓到
                    continue

                print(f"   -> [詳解] 鎖定章節標題: {curr_p.Range.Text.strip()[:30]}")
                
                start_anchor = curr_p.Range.Start
                steps = 0
                max_steps = 200 # 詳解內容可能很長
                
                while steps < max_steps:
                    steps += 1
                    try:
                        # 每次都檢查目前在 start_anchor 位置的段落
                        target_p = doc.Range(start_anchor, doc.Content.End).Paragraphs(1)
                        t_text = target_p.Range.Text.strip()
                        t_level = target_p.OutlineLevel
                        
                        # 停止條件優化：
                        # 只有當我們已經刪了標題(steps>1) 且 遇到下一個正式標題 (4.x.x) 才停止
                        if steps > 1:
                            # 如果下一段是標題，且文字不含目前的測項關鍵字，就停止
                            if t_level < 10 and len(t_text) > 0 and search_key[:8] not in t_text:
                                print(f"   -> [結束] 偵測到下一章標題: {t_text[:20]}")
                                break
                            # 絕對邊界保護
                            if any(x in t_text for x in ["Revision History", "UUT Configuration"]):
                                print(f"   -> [結束] 偵測到禁區。")
                                break
                        
                        target_p.Range.Delete()
                    except Exception as e:
                        print(f"   -> [中斷] 刪除異常: {e}")
                        break
                
                print(f"   [OK] 章節詳解清理完成。")
            else:
                keep_searching = False

    print("\n[完成] 更新並儲存...")
    doc.Fields.Update()
    doc.SaveAs(final_output, FileFormat=16)
    print(f"任務成功！檔案：{final_output}")

if __name__ == "__main__":
    remove_skipped_sections_v6()