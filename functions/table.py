def search_tables(file):
    tablescount = 0
    linktablescount = 0
    i = 1
    # Пройдите по всем таблицам в документе
    for table in file.tables[1:]:
        # Выполните некоторые действия с таблицей...
        tablescount += 1
        pass

    mask_template = """Таблица {chislo} —"""  # Задаем маску для последующей проверки по ней
    mask = mask_template.format(chislo=i)  # Добавляет переменную в маску
    for paragraph in file.paragraphs:  # Проходит по всем абзацам
        text = paragraph.text  # Записывает текст всего абзаца в переменную
        if mask in text:  # Проверяет есть ли маска в тексте
            linktablescount += 1
            i += 1
            mask = mask_template.format(chislo=i)
        text = ''
    print("Количество таблиц:", tablescount)
    print("Количество ссылок на таблицы:", linktablescount)


def search_pic(file):
    piccount = 0
    linkpiccount = 0
    j = 1
    # Проходит по всем иллюстрациям в документе
    for shape in file.inline_shapes:
        piccount += 1
        pass

    mask_template_pic = """Рисунок {chislo} —"""  # Задаем маску для последующей проверки по ней
    mask_pic = mask_template_pic.format(chislo=j)  # Добавляет переменную в маску

    for paragraph in file.paragraphs:  # Проходит по всем абзацам
        text = paragraph.text  # Записывает текст всего абзаца в переменную
        if mask_pic in text:  # Проверяет есть ли маска в тексте
            linkpiccount += 1
            j += 1
            mask_pic = mask_template_pic.format(chislo=j)
        text = ''
    print("Количество иллюстраций:", piccount)
    print("Количество ссылок на иллюстрации:", linkpiccount)

# Вывести количество таблиц
#print("Количество таблиц(все):", len(doc.tables)) #Считает все таблицы + та что на титульном

# Не работает, надо чета делать :(
# def check_links():
#     try:
#         if linktablescount != tablescount:
#             print('Кол-во таблиц и ссылок не совпадает ✕')
#         elif linktablescount == tablescount:
#             print('Кол-во таблиц и ссылок совпадает ✓')
#         elif linkpiccount != piccount:
#             print('Кол-во иллюстраций и ссылок не совпадает ✕')
#         elif linkpiccount == piccount:
#             print('Кол-во иллюстраций и ссылок совпадает ✓')
#     except Exception as exc:
#         print('Ошибка check_margin!')

#1)Кол-во табл = кол-ву ссылок (Таблица 1 - ....)
#2)Проверка "Таблца 1 -" маска
#3)Проверка ссылки на талблицу в тексте (..... в таблцие 1)

#1)Кол-во илюстрации = кол-ву ссылок (Рисунок 1 - ....)
#2)Проверка "Рисунок 1 -" маска
#3)Проверка ссылки на талблицу в тексте (..... в таблцие 1)