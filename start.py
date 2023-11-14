import docx
import functions.margin
import functions.formatting
import functions.table

import os

paths = []


def search():
    # Ищет доки
    for file in os.listdir(os.getcwd()):
        if file.endswith('.docx'):
            paths.append(file)


def names(path):
    text = 'Наименование документа: ' + path
    return text


if __name__ == '__main__':
    # выполняет все функции
    try:
        search()
        for path in paths:
            doc = docx.Document(path)
            print(names(path))
            functions.margin.margin(doc)

            functions.formatting.check_size_font(177800, doc.paragraphs)
            functions.formatting.check_name_font('Times New Roman', doc.paragraphs)

            functions.formatting.check_heading(1.25, doc.paragraphs)
            functions.formatting.check_first_line_indent(1.25, doc.paragraphs)
            functions.formatting.check_line_spacing(1.5, doc.paragraphs)
            functions.table.check_sources_merged(doc)
            functions.table.check_pic_merged(doc)
            functions.table.check_table_merged(doc)
            doc.save(path)
    except Exception as e:
        print('Ошибка!')