我有一份 Word ，其中有一個表格，有辦法用 Python ，去增加新的一橫列嗎？



如果可以該怎麼做呢？

這是一個非常實用的自動化需求！在 Python 中，處理 Word 檔案（.docx）最標準且強大的函式庫是 python-docx。

你可以透過這套工具輕鬆地開啟現有的 Word 檔、定位到特定的表格，並增加新的橫列（Row）。

實作步驟
安裝函式庫：
如果你還沒安裝，請先在終端機執行：

Bash
pip install python-docx
程式碼範例：
假設你的 Word 檔名為 test.docx，且裡面已經有一個表格。

Python
from docx import Document

# 1. 載入現有的 Word 檔案
doc = Document('test.docx')

# 2. 選取表格 (索引從 0 開始，第一個表格就是 tables[0])
if doc.tables:
    table = doc.tables[0]
    
    # 3. 增加一個新橫列
    new_row = table.add_row()
    
    # 4. 為新橫列的儲存格填入資料
    # 假設表格有三欄，cells 的索引分別是 0, 1, 2
    new_row.cells[0].text = "新數據 A"
    new_row.cells[1].text = "新數據 B"
    new_row.cells[2].text = "新數據 C"

    # 5. 儲存變更 (可以覆蓋原檔或另存新檔)
    doc.save('updated_test.docx')
    print("成功增加橫列並儲存檔案！")
else:
    print("文件中找不到任何表格。")
重點筆記
定位表格：doc.tables 是一個清單。如果你的 Word 裡有多個表格，你需要確認你要修改的是哪一個。例如，第二個表格就是 doc.tables[1]。

樣式繼承：使用 add_row() 增加的列，通常會自動繼承前一列的格式（如字體大小、對齊方式等）。

存取儲存格：你也可以透過 table.cell(row_index, col_index) 來精確定位某個儲存格進行修改。

大量資料：如果你需要從 Excel 或資料庫匯入大量資料到 Word 表格，可以用 for 迴圈搭配 add_row() 來快速生成。

如果你在執行時遇到表格內有「合併儲存格」的情況，處理起來會稍微複雜一點，有需要的話可以再跟我說！

如果要看這個 Word 中，有多少表格，應該怎麼做？

要查看 Word 檔案中有多少個表格，同樣可以使用 python-docx 函式庫。

你可以透過 len(doc.tables) 來直接取得表格的總數。以下是完整的實作程式碼：

程式碼範例
Python
from docx import Document

# 1. 載入 Word 檔案
doc = Document('test.docx')

# 2. 取得表格清單
all_tables = doc.tables

# 3. 印出表格總數
table_count = len(all_tables)
print(f"這份文件中共有 {table_count} 個表格。")

# 4. (進階) 顯示每個表格分別有多少欄與列
for i, table in enumerate(all_tables):
    rows = len(table.rows)
    cols = len(table.columns)
    print(f"表格 {i}：共有 {rows} 列, {cols} 欄")
常見問題補充
如果結果是 0：請確認你的檔案格式是 .docx。如果是舊版的 .doc 檔案，python-docx 無法直接讀取，需要先在 Word 中另存為 .docx。

巢狀表格（表格中還有表格）：上述的方法主要是計算「主體文件」中的表格。如果你的表格是畫在一個「文字方塊」或是「頁首頁尾」裡面，則需要另外定位到該區塊才能計算。

快速檢查：如果你只是想在開發時快速確認索引，enumerate（如範例第 12 行）非常好用，它可以幫你標記出每個表格的索引號碼（0, 1, 2...），方便你後續指定操作對象。

請問 表格 會被分類在不同 大綱 段落 中嗎？或是一個 Word 檔中的表格，就是流水編號從 ( 0, 1 , 2... )？

這是一個非常好的問題！這牽涉到 Word 檔案底層的結構。

簡單來說：在 python-docx 的邏輯中，表格是「獨立」於段落（Paragraph）之外的。

1. 它是如何編號的？
doc.tables 確實是採用流水編號 (0, 1, 2...)。
無論你的表格是在第一頁的標題下，還是最後一頁的附錄中，python-docx 會由上而下掃描整個檔案，將所有找到的表格依序放入 doc.tables 這個清單裡。

優點：你可以非常快速地統計總數，或抓取第一個、最後一個表格。

缺點：如果表格很多，你很難只從索引編號（例如 tables[5]）判斷它到底屬於哪個大綱章節。

2. 表格會被分類在「段落」中嗎？
不會。 在 Word 的物件模型裡，文件內容主要分為兩類：

Paragraphs (段落)：純文字、大綱標題等。

Tables (表格)：獨立的結構體。

當你執行 doc.paragraphs 時，裡面完全不會包含表格內容；反之，doc.tables 裡也完全不屬於任何段落。

3. 進階：如何知道表格在哪個「大綱」下？
如果你需要知道某個表格是在哪個標題（例如「第二章」）之後出現的，不能直接用 doc.tables，而必須遍歷文件的 「主體元件」(Body Elements)。

你可以使用以下邏輯來判斷表格的位置：

Python
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document('test.docx')

current_heading = "未分類（文件開頭）"

# 遍歷文件中的所有頂層元件 (段落或表格)
for element in doc.element.body:
    # 如果是段落
    if element.tag.endswith('p'):
        from docx.text.paragraph import Paragraph
        para = Paragraph(element, doc)
        
        # 檢查是否為標題 (Heading 1, 2, 3...)
        if para.style.name.startswith('Heading'):
            current_heading = para.text
            
    # 如果是表格
    elif element.tag.endswith('tbl'):
        print(f"發現表格！它位於標題「{current_heading}」之下。")
總結
預設行為：表格就是簡單的 0, 1, 2 流水號，不分章節。

關聯章節：如果你要「按大綱分類」表格，必須手動寫程式去追蹤在表格出現之前，最後一個出現的標題是什麼。