def margin(file):
    # выводит поля
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
    # проверяет поля
    try:
        if round(margin.cm, 1) == expected_value:
            print(f'{name} поле: {round(margin.cm, 1)}см ✓')
        else:
            print(f'{name} поле: {round(margin.cm, 1)} см ✕ ({name} поле должно иметь границу {expected_value}см!)')
    except Exception as exc:
        print('Ошибка check_margin!')