import win32com.client as win32
import os
import pandas as pd

def interactive_tagging(file_path):
    # 1. 啟動 Word 並開啟檔案
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True  # 讓你一邊操作一邊看 Word 的變化
    
    abs_path = os.path.abspath(file_path)
    doc = word.Documents.Open(abs_path)
    
    mapping_data = []
    print(f"--- 進入互動模式：正在處理 {os.path.basename(file_path)} ---")
    print("輸入 'exit' 結束並存檔\n")

    try:
        while True:
            old_text = input("請輸入 Word 中要被取代的『舊字串』(例如 HA2026-HC): ").strip()
            if old_text.lower() == 'exit':
                break
            
            tag_name = input(f"請輸入要替換成的『標籤名稱』(例如 PROJECT_NAME): ").strip()
            new_tag = f"{{{{ {tag_name} }}}}"
            
            # 2. 執行 Word 取代功能 (遍歷所有 StoryRanges 以確保包含內文與頁首頁尾)
            found = False
            for story in doc.StoryRanges:
                # Execute 參數順序: FindText, MatchCase, MatchWholeWord, MatchWildcards, ...
                if story.Find.Execute(old_text, False, False, False, False, False, True, 1, True, new_tag, 2):
                    found = True
                
                # 處理可能的後續 Story (例如多個節的頁首頁尾)
                next_story = story.NextStoryRange
                while next_story:
                    if next_story.Find.Execute(old_text, False, False, False, False, False, True, 1, True, new_tag, 2):
                        found = True
                    next_story = next_story.NextStoryRange

            
            if found:
                print(f"成功：已將 '{old_text}' 替換為 '{new_tag}'")
                mapping_data.append({"Tag_Name": tag_name, "Value": old_text})
            else:
                print(f"找不到字串: '{old_text}'，請重新輸入。")

            cont = input("\n是否繼續下一個標籤替換？(Y/n): ").strip().lower()
            if cont == 'n':
                break

    finally:
        # 3. 儲存結果
        # 另存 Word 模板
        template_name = os.path.splitext(abs_path)[0] + "_Template.docx"
        doc.SaveAs(template_name, FileFormat=16) # 16 是 .docx 格式
        print(f"\n[系統] 模板已儲存至: {template_name}")
        
        # 儲存 Excel 對照表
        if mapping_data:
            df = pd.DataFrame(mapping_data)
            excel_name = "data.xlsx"
            df.to_excel(excel_name, index=False, sheet_name='SYS_CONF')
            print(f"[系統] 對照表已紀錄至: {excel_name}")
        
        doc.Close()
        word.Quit()

if __name__ == "__main__":
    target_file = "RD260309A09-PSG_HA2026-HC_DVT Test Report Ver A.DOC"
    if os.path.exists(target_file):
        interactive_tagging(target_file)
    else:
        print(f"找不到檔案: {target_file}")
