import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

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


def margin(file):
    try:
        sections = file.sections
        print('Поля:')
        for section in sections:
            check_margin(section.top_margin, 2.0, 'Верхнее')
            check_margin(section.bottom_margin, 2.0, 'Нижнее')
            check_margin(section.left_margin, 3.0, 'Левое')
            check_margin(section.right_margin, 1.5, 'Правое')
        print('\n', end='')
    except Exception as exc:
        print('Ошибка margin!')


def check_margin(margin, expected_value, name):
    try:
        if round(margin.cm, 1) == expected_value:
            print(f'{name} поле: {round(margin.cm, 1)}см ✓')
        else:
            print(f'{name} поле: {round(margin.cm, 1)} см ✕ ({name} поле должно иметь границу {expected_value}см!)')
    except Exception as exc:
        print('Ошибка check_margin!')


def text_formatting(file):
    try:
        print('Отступы первой строки:', end='')
        paragraphs = file.paragraphs
        for paragraph in paragraphs:
            formatting = paragraph.paragraph_format
            check_text_formatting(formatting.first_line_indent, paragraph.alignment, paragraph.text, 1.25)
        print('\n')
    except Exception as exc:
        print('Ошибка text_formatting!')


def check_text_formatting(indent, alignment, text, expected_value_first_line):
    try:
        if indent != 0 and alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            print(f'{text} ✕ (Отступ первой строки заголовка должен быть 0см!)')
        elif round(indent.cm, 2) == expected_value_first_line or text == '':
            return 0
        elif indent == 0 and alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            return 0
        else:
            print(f'{text} ✕ (Отступ первой строки абзаца должен быть {expected_value_first_line}см!)')
    except Exception as exc:
        print('Ошибка check_text_formatting!')


if __name__ == '__main__':
    try:
        search()
        for path in paths:
            doc = docx.Document(path)
            properties(doc, path)
            margin(doc)
            text_formatting(doc)
    except Exception as exc:
        print('Ошибка!')
