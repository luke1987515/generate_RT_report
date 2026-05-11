from docxtpl import DocxTemplate
import pandas as pd
import os

def generate_styled_report(excel_path, template_path, output_path):
    # 1. 讀取 Excel (假設第一欄是標籤，第二欄是填寫內容)
    df = pd.read_excel(excel_path)
    
    # 2. 轉換為字典格式
    # 我們將第一欄 (index 0) 當作 Key，第二欄 (index 1) 當作 Value
    # docxtpl 不需要我們手動寫 replace，它會自動對應
    context = dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1].astype(str)))
    

    
    # 3. 載入 Word 模板
    doc = DocxTemplate(template_path)

    # 在 doc.render(context) 之前加入
    print("目前抓到的取代字典內容：")
    print(context)

    # 在 doc.render(context) 之前，手動插一行測試
    context['{{RESULT}}'] = "TEST_DISPLAY"
    
    # 4. 渲染 (這一步會處理內文、表格、頁首頁尾，並保留格式)
    doc.render(context)
    
    # 5. 儲存檔案
    doc.save(output_path)
    print(f"✨ 報告產出成功！格式已完美保留：{output_path}")

# 執行設定
EXCEL_FILE = "data_input.xlsx"
TEMPLATE_FILE = "template.docx"
OUTPUT_FILE = "final_styled_report.docx"

generate_styled_report(EXCEL_FILE, TEMPLATE_FILE, OUTPUT_FILE)