from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_BREAK


def text_formatting(file):
    # выводит все функции
    try:
        paragraphs = file.paragraphs
        for paragraph in paragraphs:
            formatting = paragraph.paragraph_format
            for run in paragraph.runs:
                check_name_font('Times New Roman', paragraph, run)
                check_size_font(14.0, paragraph, run)
                check_heading(paragraph.alignment, paragraph, run)
                # check_subheading(paragraph.alignment, paragraph)
                check_first_line_indent(formatting.first_line_indent, paragraph.alignment, 1.25, paragraph, run)
                check_line_spacing(formatting.line_spacing, 1.5, paragraph, run)
        print('\n')
    except Exception as exc:
        print('Ошибка text_formatting!')


def check_name_font(expected_value, paragraph, run):
    # чекает шрифт
    try:
        if paragraph.text != '' or run.text != ' ':
            # если абзац не пустой
            if run.font.name is not None:
                # тут по другому не сделать, т.к. шрифт Times New Roman он принимает за None
                if run.font.name != expected_value:
                    # на всякий случай сделал еще это
                    comment = paragraph.add_comment('Шрифт должен быть Times New Roman!')
                    comment.author = 'bot'
    except Exception as exc:
        print('Ошибка check_name_font!')


def check_size_font(expected_value, paragraph, run):
    # проверка размера шрифта
    try:
        if round(run.font.size.pt, 1) != expected_value:
            comment = paragraph.add_comment('Размер шрифта должен быть 14пт!')
            comment.author = 'bot'
    except Exception as exc:
        print('Ошибка check_size_font!')


def check_first_line_indent(indent, alignment, expected_value_first_line, paragraph, run):
    # проверка отсутпа первого абзаца (вот тут трудности)
    try:
        if indent.cm == 0.0 and alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            # Если отступ равен 0 и текст по центру
            return 0
        elif indent.cm != 0.0 and alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            comment = paragraph.add_comment(f'Отступ строки должен быть 0см!')
            comment.author = 'bot'
        elif round(indent.cm, 2) == expected_value_first_line and alignment != WD_PARAGRAPH_ALIGNMENT.CENTER:
            # Если отступ равен введённому знаичению
            return 0
        elif round(indent.cm, 2) != expected_value_first_line and alignment != WD_PARAGRAPH_ALIGNMENT.CENTER:
            comment = paragraph.add_comment(f'Отступ первой строки абзаца должен быть {expected_value_first_line}см!')
            comment.author = 'bot'
    except Exception as exc:
        print('Ошибка check_first_line_indent!')


def check_line_spacing(spacing, expected_value_line_spacing, paragraph, run):
    # межстрочный интервал, обычно он комментит по 3-6 строк на одну строку
    try:
        if paragraph.text != '' or run.text != ' ':
            if spacing is not None:
                if round(spacing, 1) != expected_value_line_spacing:
                    comment = paragraph.add_comment(f'Межстрочный интервал должен быть {expected_value_line_spacing}см!')
                    comment.author = 'bot'
    except Exception as exc:
        print('Ошибка check_line_spacing!')


def check_heading(alignment, paragraph, run):
    # проверк азаголовка (надо улучшить), потому что за заголов он принимает все, чо пишется в начале жирным
    try:
        if paragraph.text != '' or run.text != ' ':
            if run.text.lower() != run.text and run.bold and alignment != WD_PARAGRAPH_ALIGNMENT.CENTER:
                comment = paragraph.add_comment(f'Выравнивание должно быть по центру!')
                comment.author = 'bot'
    except Exception as exc:
        print('Ошибка check_heading!')


# из этих важна только функция подхаголовка, но надо закончить хотя бы с заголовком

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
