from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_BREAK

def text_formatting(file):
    try:
        paragraphs = file.paragraphs
        for paragraph in paragraphs:
            formatting = paragraph.paragraph_format
            # check_name_font('Times New Roman', paragraph)
            # check_size_font(14.0, paragraph)
            check_heading(paragraph.alignment, paragraph)
            # check_subheading(paragraph.alignment, paragraph)
            # check_first_line_indent(formatting.first_line_indent, paragraph.alignment, paragraph.text, 1.25, paragraph)
            # check_line_spacing(formatting.line_spacing, 1.25, paragraph)
            # check_space_after(formatting.space_after, paragraph.alignment, 8.0, paragraph)
            # check_space_before(formatting.space_before,0, paragraph)
        print('\n')
    except Exception as exc:
        print('Ошибка text_formatting!')


# def check_name_font(expected_value, paragraph):
#     name = paragraph.style.font.name
#     if name != expected_value:
#         comment = paragraph.add_comment('Шрифта должен быть Times New Roman!')
#         comment.author = 'bot'


def check_size_font(expected_value, paragraph):
    try:
        for run in paragraph.runs:
            size = run.font.size.pt
            # if name is not None or name != '':
            #     comment = paragraph.add_comment('Шрифт должен быть Times New Roman!')
            #     comment.author = 'bot'
            if size != expected_value:
                comment = paragraph.add_comment('Размер шрифта должен быть 14пт!')
                comment.author = 'bot'
    except Exception as exc:
        print('Ошибка check_size_font!')


def check_first_line_indent(indent, alignment, text, expected_value_first_line, paragraph):
    try:
        if indent != 0 and alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            # Если отступ не равен 0 и текст по центру
            comment = paragraph.add_comment('Отступ первой строки заголовка должен быть 0см!')
            comment.author = 'bot'
        elif round(indent.cm, 2) == expected_value_first_line or text == '':
            # Если отступ равен введённому знаичению или текст отсутсвует
            return 0
        elif indent == 0 and alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            # Если отступ равен 0 знаичению и текст по центру
            return 0
        else:
            comment = paragraph.add_comment(f'Отступ первой строки абзаца должен быть {expected_value_first_line}см!')
            comment.author = 'bot'
    except Exception as exc:
        print('Ошибка check_first_line_indent!')


def check_line_spacing(spacing, expected_value_line_spacing, paragraph):
    try:
        if spacing is not None:
            comment = paragraph.add_comment(f'Межстрочный интервал должен быть {expected_value_line_spacing}см!')
            comment.author = 'bot'
    except Exception as exc:
        print('Ошибка check_line_spacing!')


def check_heading(alignment, paragraph):
    try:
        for run in paragraph.runs:
            if run.text.lower() != run.text and run.bold and alignment != WD_PARAGRAPH_ALIGNMENT.CENTER:
                comment = paragraph.add_comment(f'Выравнивание должно быть по центру!')
                comment.author = 'bot'
    except Exception as exc:
        print('Ошибка check_heading!')


# def check_subheading(alignment, paragraph):
#     try:
#         for run in paragraph.runs:
#             if run.text[0].upper() and run.bold and alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
#                 comment = paragraph.add_comment(f'Выравнивание должно быть по центру!')
#                 comment.author = 'bot'
#     except Exception as exc:
#         print('Ошибка check_subheading!')


# def check_space_after(spacing, alignment, expected_value_space_after, paragraph):
#     try:
#         print(spacing)
#         if spacing is not None:
#             if round(spacing.pt, 1) != expected_value_space_after and alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
#                 comment = paragraph.add_comment(f'Интервал после строки заголовка должен быть {expected_value_space_after}пт!')
#                 comment.author = 'bot'
#             elif round(spacing.pt, 1) != 0 and alignment != WD_PARAGRAPH_ALIGNMENT.CENTER:
#                 comment = paragraph.add_comment(f'Интервал после строки должен быть 0пт!')
#                 comment.author = 'bot'
#     except Exception as exc:
#         print('Ошибка check_space_after!')


# def check_space_before(spacing, expected_value_space_before, paragraph):
#     try:
#         if spacing is not None:
#             comment = paragraph.add_comment(f'Интервал перед строкой должен быть {expected_value_space_before}см!')
#             comment.author = 'bot'
#     except Exception as exc:
#         print('Ошибка check_space_before!')
