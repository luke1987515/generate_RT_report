from docx import Document

def replace_text_in_word(file_path, old_text, new_text, output_path):
    doc = Document(file_path)

    # 1. 替換段落中的文字
    for paragraph in doc.paragraphs:
        if old_text in paragraph.text:
            paragraph.text = paragraph.text.replace(old_text, new_text)

    # 2. 替換表格中的文字
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if old_text in paragraph.text:
                        paragraph.text = paragraph.text.replace(old_text, new_text)

    # 3. 替換頁首與頁尾 (New!)
    for section in doc.sections:
        # 處理頁首
        header = section.header
        for paragraph in header.paragraphs:
            if old_text in paragraph.text:
                paragraph.text = paragraph.text.replace(old_text, new_text)
        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if old_text in paragraph.text:
                            paragraph.text = paragraph.text.replace(old_text, new_text)

        # 處理頁尾
        footer = section.footer
        for paragraph in footer.paragraphs:
            if old_text in paragraph.text:
                paragraph.text = paragraph.text.replace(old_text, new_text)
        for table in footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if old_text in paragraph.text:
                            paragraph.text = paragraph.text.replace(old_text, new_text)

    doc.save(output_path)
    print(f"包含頁首頁尾替換完成！已儲存至: {output_path}")

# 執行
replace_text_in_word("原始文件.docx", "舊字串", "新字串", "修改後文件.docx")