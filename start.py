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


def properties(file, path):
    # настройки дока
    properties = file.core_properties
    print('Наименование документа:', path)


if __name__ == '__main__':
    # выполняет все функции
    try:
        search()
        for path in paths:
            doc = docx.Document(path)
            properties(doc, path)
            functions.margin.margin(doc)
            functions.formatting.text_formatting(doc)
            functions.table.search_tables(doc)
            functions.table.search_pic(doc)
            doc.save(path)
    except Exception as e:
        print('Ошибка!')

