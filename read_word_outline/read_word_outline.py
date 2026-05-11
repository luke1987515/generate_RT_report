import win32com.client as win32
import os
import sys
import io

# 強制設定輸出編碼為 UTF-8，解決 Windows 終端機編碼問題
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def read_word_headings(file_path):
    # 轉換為絕對路徑，避免 win32com 找不到檔案
    abs_path = os.path.abspath(file_path)
    
    if not os.path.exists(abs_path):
        print(f"錯誤: 找不到檔案 {abs_path}")
        return

    # 啟動 Word 應用程式
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
    except AttributeError:
        # 有時候 EnsureDispatch 會失敗，嘗試 Dispatch
        word = win32.Dispatch('Word.Application')
        
    word.Visible = False  # 在後台執行
    
    try:
        # 開啟文件
        doc = word.Documents.Open(abs_path)
        
        print(f"正在讀取檔案：{os.path.basename(file_path)}\n")
        print("--- 文件大綱架構 ---")
        
        found_any = False
        for para in doc.Paragraphs:
            # 取得段落樣式名稱
            style_name = para.Style.NameLocal
            
            # 跳過目錄相關的項目 (TOC 1, TOC 2, 目錄 1 等)
            if "TOC" in style_name or "目錄" in style_name:
                continue

            # 取得大綱層級 (1-9 為標題，10 為內文)
            # wdOutlineLevel1 = 1, wdOutlineLevel9 = 9, wdOutlineLevelBodyText = 10
            try:
                level = para.OutlineLevel
            except Exception:
                level = 10 # 預設為內文
            
            if level <= 9:
                found_any = True
                indent = "  " * (level - 1)
                text = para.Range.Text.strip()
                if text: # 忽略空的標題段落
                    print(f"{indent}L{level}: {text}")
            else:
                # 備案：如果 OutlineLevel 不準確，檢查樣式名稱
                if "標題" in style_name or "Heading" in style_name:
                    found_any = True
                    level_str = style_name.replace("標題", "").replace("Heading", "").strip()
                    if level_str.isdigit():
                        l = int(level_str)
                        indent = "  " * (l - 1)
                        print(f"{indent}L{l} (Style): {para.Range.Text.strip()}")
        
        if not found_any:
            print("未偵測到任何標題架構。")
            print("\n調試資訊：前 10 個段落的樣式與層級：")
            for i, para in enumerate(doc.Paragraphs):
                if i >= 10: break
                print(f"段落 {i+1}: Style={para.Style.NameLocal}, OutlineLevel={para.OutlineLevel}, Text={para.Range.Text.strip()[:20]}...")

        doc.Close()
    except Exception as e:
        print(f"發生錯誤: {e}")
    finally:
        # 關閉 Word 程式
        word.Quit()

# 使用範例
if __name__ == "__main__":
    file_name = "Template.doc" 
    read_word_headings(file_name)