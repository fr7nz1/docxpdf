import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_BREAK

doc = docx.Document('1.docx')


def check_num_tables(file):
    tablescount = 0
    # Проходим по всем таблицам в документе
    for table in doc.tables[1:]:
        tablescount += 1
        pass
    # print("Кол-во таблиц:", tablescount)
    return tablescount


def check_link_tables(file, tablescount):
    linktablescount = 0
    i = 1
    mask_template = """Таблица {chislo} —"""  # Задаем маску для последующей проверки по ней
    mask = mask_template.format(chislo=i)  # Добавляет переменную в маску
    for paragraph in doc.paragraphs:  # Проходит по всем абзацам
        if mask in paragraph.text:  # Проверяет есть ли маска в тексте
            linktablescount += 1
            i += 1
            mask = mask_template.format(chislo=i)
            # if linktablescount >= tablescount:
                # тут мы делаем пометку в данном параграфе что ссылок стало больше чем таблиц
                # comment = paragraph.add_comment("Кол-во ссылок на таблицы и кол-во таблиц не совпадает!")
                # comment.author = 'bot'
    # print("Кол-во ссылок на таблицы:", linktablescount)
    return linktablescount


def check_num_pic(file):
    piccount = 0
    # Проходит по всем иллюстрациям в документе
    for shape in doc.inline_shapes:
        piccount += 1
        pass
    # print("Кол-во иллюстраций:", piccount)
    return piccount


def check_link_pic(file):
    linkpiccount = 0
    i = 1
    mask_template_pic = """Рисунок {chislo} —"""  # Задаем маску для последующей проверки по ней
    mask_pic = mask_template_pic.format(chislo=i)  # Добавляет переменную в маску
    for paragraph in doc.paragraphs:  # Проходит по всем абзацам
        if mask_pic in paragraph.text:  # Проверяет есть ли маска в тексте
            linkpiccount += 1
            i += 1
            mask_pic = mask_template_pic.format(chislo=i)
    # print("Кол-во ссылок на иллюстрации:", linkpiccount)
    return linkpiccount


def check_num_sources(file):
    sourcescount = 0
    i = 1
    mask_template_sources = """{chislo}. """
    mask_sources = mask_template_sources.format(chislo=i)

    for paragraph in doc.paragraphs:  # Нужно придумать как начать сразу с последней страницы
        if mask_sources in paragraph.text:
            sourcescount += 1
            i += 1
            mask_sources = mask_template_sources.format(chislo=i)
    # print("Кол-во источников:", sourcescount)
    return sourcescount


def check_link_sources(file):
    linksourcecount = 0
    i = 1
    mask_template_link_sources = """[{chislo}]"""
    mask_link_sources = mask_template_link_sources.format(chislo=i)

    for paragraph in doc.paragraphs:  # Нужно придумать как начать сразу с последней страницы
        if mask_link_sources in paragraph.text:  # Проверяет есть ли маска в тексте
            linksourcecount += 1
            i += 1
            mask_link_sources = mask_template_link_sources.format(chislo=i)
    # print("Кол-во ссылок на источники:", linksourcecount)
    return linksourcecount


def check_sources_merged(file):
    sourcescount = 0
    linksourcecount = 0
    i = 0
    j = 0
    mask_template_sources = """{chislo}. """
    mask_sources = mask_template_sources.format(chislo=i)
    mask_template_link_sources = """[{chislo}]"""
    mask_link_sources = mask_template_link_sources.format(chislo=j)

    mask_title = """БИБЛИОГРАФИЧЕСКИЙ СПИСОК"""

    for paragraph in doc.paragraphs:
        if mask_sources in paragraph.text:
            sourcescount += 1
            i += 1
            mask_sources = mask_template_sources.format(chislo=i)
        elif mask_link_sources in paragraph.text:
            linksourcecount += 1
            j += 1
            mask_link_sources = mask_template_link_sources.format(chislo=j)

    if sourcescount != linksourcecount:
        for paragraph in doc.paragraphs:
            if mask_title in paragraph.text:
                comment = paragraph.add_comment('Кол-во источников и кол-во ссылок на источники не совпадает!')
                comment.author = 'bot'
                break


def check_pic_merged(file):
    piccount = 0
    linkpiccount = 0
    i = 1
    mask_template_pic = """Рисунок {chislo} —"""  # Задаем маску для последующей проверки по ней
    mask_pic = mask_template_pic.format(chislo=i)  # Добавляет переменную в маску
    # Проходит по всем иллюстрациям в документе
    for shape in doc.inline_shapes:
        piccount += 1
        pass
    for paragraph in doc.paragraphs:  # Проходит по всем абзацам
        if mask_pic in paragraph.text:  # Проверяет есть ли маска в тексте
            linkpiccount += 1
            i += 1
            mask_pic = mask_template_pic.format(chislo=i)

    if piccount != linkpiccount: # пишем так, чтоб вывод был на поля 1 страницы
        for paragraph in doc.paragraphs:
            comment = paragraph.add_comment('Кол-во иллюстраций и кол-во ссылок на иллюстрации не совпадает!')
            comment.author = 'bot'
            break


def check_table_merged(file):
    tablescount = 0
    linktablescount = 0
    i = 1
    mask_template = """Таблица {chislo} —"""  # Задаем маску для последующей проверки по ней
    mask = mask_template.format(chislo=i)  # Добавляет переменную в маску
    # Проходим по всем таблицам в документе
    for table in doc.tables[1:]:
        tablescount += 1
        pass
    for paragraph in doc.paragraphs:  # Проходит по всем абзацам
        if mask in paragraph.text:  # Проверяет есть ли маска в тексте
            linktablescount += 1
            i += 1
            mask = mask_template.format(chislo=i)

    if tablescount != linktablescount: # пишем так, чтоб вывод был на поля 1 страницы
        for paragraph in doc.paragraphs:
            comment = paragraph.add_comment('Кол-во таблиц и кол-во ссылок на таблицы не совпадает!')
            comment.author = 'bot'
            break


# def check_links(tablescount, linktablescount, piccount, linkpiccount, sourcescount, linksourcecount):
#     try:
#         if linksourcecount != sourcescount:
#             print('Кол-во ссылок на источники и кол-во источников не совпадает ✕')
#         elif linktablescount == tablescount:
#             print('Кол-во таблиц и ссылок совпадает ✓')
#         elif linkpiccount != piccount:
#             print('Кол-во иллюстраций и ссылок не совпадает ✕')
#         elif linkpiccount == piccount:
#             print('Кол-во иллюстраций и ссылок совпадает ✓')
#     except Exception as exc:
#         print('Ошибка check_margin!')


# print("Количество таблиц(все):", len(doc.tables)) #Считает все таблицы + та что на титульном
# check_links(check_link_sources(doc), check_num_sources(doc), check_link_pic(doc), check_num_pic(doc))

# check_sources_merged(doc)
# check_num_pic(doc)
# check_link_pic(doc)
# check_num_sources(doc)
# check_link_sources(doc)