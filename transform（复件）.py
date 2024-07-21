import csv
from docx import Document


def docx_to_csv(input_file, output_file):
    # 打开Word文档
    doc = Document(input_file)

    # 创建CSV写入器
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)

        # 遍历文档中的所有表格
        for table in doc.tables:
            # 遍历表格中的每一行
            for row in table.rows:
                # 提取单元格文本并写入CSV
                writer.writerow([cell.text for cell in row.cells])


# 调用函数转换文档
docx_to_csv('表格3.docx', 'output3.csv')