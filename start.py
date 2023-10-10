import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

import os

paths = []


def searchDocx():
    for file in os.listdir(os.getcwd()):
        if file.endswith('.docx'):
            paths.append(file)


def propertiesDocx(file, path):
    properties = file.core_properties
    print('Наименование документа:', path)
    print('Автор документа:', properties.author)
    print('Дата и время создания документа:', properties.created, '\n')


def marginDocx(file):
    sections = file.sections
    print('Поля:')
    for section in sections:
        check_margin(section.top_margin, 2.0, 'Верхнее')
        check_margin(section.bottom_margin, 2.0, 'Нижнее')
        check_margin(section.left_margin, 3.0, 'Левое')
        check_margin(section.right_margin, 1.5, 'Правое')
    print('\n')

#def text_formatting(file):
#    paragraphs = file.paragraphs
#    for paragraph in paragraphs:
#        formatting = paragraph.paragraph_format


def check_text_formatting(file):
    print('Отступы первой строки:')
    for paragraph in file.paragraphs:
        formatting = paragraph.paragraph_format
        if formatting.first_line_indent != 0 and paragraph.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            print(paragraph.text, ' ✕ (Отступ первой строки заголовка должен быть 0см!)')
        elif round(formatting.first_line_indent.cm, 2) == 1.25 or paragraph.text == '':
            continue
        elif formatting.first_line_indent == 0 and paragraph.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            continue
        else:
            print(paragraph.text, ' ✕ (Отступ первой строки абзаца должен быть 1.25см!)', formatting.space_before.pt)
    print('\n')


        # print('Отступ после абзаца:', formatting.space_after)
        # print('Отступ слева:', formatting.left_indent)
        # print('Отступ справа:', formatting.right_indent)


def check_margin(margin, expected_value, name):
    try:
        if round(margin.cm, 1) == expected_value:
            print(f'{name} поле: {round(margin.cm, 1)}см ✓')
        else:
            print(f'{name} поле: {round(margin.cm, 1)} см ✕ ({name} поле должно иметь границу {expected_value}см!)')
    except Exception as exc:
        print('Ошибка check_margin!')


if __name__ == '__main__':
    try:
        searchDocx()
        for path in paths:
            doc = docx.Document(path)
            propertiesDocx(doc, path)
            marginDocx(doc)
            # for paragraph in doc.paragraphs:
            #     print(paragraph.alignment)
            check_text_formatting(doc)
    except Exception as exc:
        print('Ошибка!')
