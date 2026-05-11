import pandas as pd
from docx import Document
import os

def replace_from_excel(excel_path, template_path, output_path):
    # 1. 讀取 Excel
    # 假設沒有標題列，或是第一列就是資料，可用 header=None
    # 如果第一列是「標籤、內容」，則不需 header=None
    df = pd.read_excel(excel_path)
    
    # 2. 將 Excel 轉成字典 (假設第一欄是 Key, 第二欄是 Value)
    # iloc[:, 0] 表示所有列的第一欄，iloc[:, 1] 表示第二欄
    replacements = dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1].astype(str)))
    
    # 3. 載入 Word 模板
    doc = Document(template_path)

    # 定義替換邏輯 (段落與表格)
    def do_replace(paragraphs):
        for p in paragraphs:
            for key, val in replacements.items():
                if key in p.text:
                    p.text = p.text.replace(key, val)

    def do_replace_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    do_replace(cell.paragraphs)

    # 4. 執行替換 (內文、頁首、頁尾)
    do_replace(doc.paragraphs)
    do_replace_tables(doc.tables)
    
    for section in doc.sections:
        do_replace(section.header.paragraphs)
        do_replace_tables(section.header.tables)
        do_replace(section.footer.paragraphs)
        do_replace_tables(section.footer.tables)

    # 5. 儲存
    doc.save(output_path)
    print(f"✅ 已根據 {excel_path} 完成替換！產出檔案：{output_path}")

# 使用範例
replace_from_excel("data_input.xlsx", "template.docx", "final_report.docx")