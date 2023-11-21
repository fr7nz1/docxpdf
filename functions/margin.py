def margin(file):
    # выводит поля
    try:
        # берём поля в докменте
        sections = file.sections
        # проходим по ним и проверяем
        for section in sections:
            check_margin(file, section.top_margin, 2.0, 'Верхнее')
            check_margin(file, section.bottom_margin, 2.0, 'Нижнее')
            check_margin(file, section.left_margin, 3.0, 'Левое')
            check_margin(file, section.right_margin, 1.5, 'Правое')
    except Exception as exc:
        print('Ошибка margin!')


def check_margin(file, margin, expected_value, name):
    # проверяет поля
    try:
        # если поле не равен введенному занчению
        if round(margin.cm, 1) != expected_value:
            # написание комментав первой строке
            for paragraph in file.paragraphs:
                comment = paragraph.add_comment(f'{name} поле должно иметь границу {expected_value}см!)')
                comment.author = 'bot'
                break
    except Exception as exc:
        print('Ошибка check_margin!')