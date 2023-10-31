import docx
import functions.margin
import functions.formatting

import os

paths = []


def search():
    for file in os.listdir(os.getcwd()):
        if file.endswith('.docx'):
            paths.append(file)


def properties(file, path):
    properties = file.core_properties
    print('Наименование документа:', path)
    print('Автор документа:', properties.author)
    print('Дата и время создания документа:', properties.created, '\n')


if __name__ == '__main__':
    try:
        search()
        for path in paths:
            doc = docx.Document(path)
            properties(doc, path)
            functions.margin.margin(doc)
            functions.formatting.text_formatting(doc)
            doc.save(path)
    except Exception as e:
        print('Ошибка!')

