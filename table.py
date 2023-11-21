import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_BREAK

# doc = docx.Document('1.docx')
doc = docx.Document('Тест.docx')
# doc = docx.Document('Тест(табл).docx')


def check_num_tables(file):
    tablescount = 0
    # Проходим по всем таблицам в документе
    for table in doc.tables[1:]:
        tablescount += 1
        pass
    # print("Кол-во таблиц:", tablescount)
    return tablescount


def check_link_tables(file):
    linktablescount = 0
    i = 1
    mask_template = """Таблица {chislo} –"""  # Задаем маску для последующей проверки по ней
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
    mask_template_pic = """Рисунок {chislo} –"""  # Задаем маску для последующей проверки по ней
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

    flag = False
    mask_title = """Список литературы"""

    for paragraph in doc.paragraphs:
        if mask_title in paragraph.text:
            flag = True

        if flag == True:
            if paragraph.style.name == 'List Paragraph':
                sourcescount += 1
                print(paragraph.text)
    # i = 1
    # mask_template_sources = """{chislo}. """
    # mask_sources = mask_template_sources.format(chislo=i)
    #
    # for paragraph in doc.paragraphs:  # Нужно придумать как начать сразу с последней страницы
    #     if mask_sources in paragraph.text:
    #         sourcescount += 1
    #         i += 1
    #         mask_sources = mask_template_sources.format(chislo=i)
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

    flag = False
    mask_title = """Список литературы"""  # По госту должен быть БИБЛИОГРАФИЧЕСКИЙ СПИСОК
    # mask_title = """БИБЛИОГРАФИЧЕСКИЙ СПИСОК"""

    j = 1
    mask_template_link_sources = """[{chislo}]"""
    mask_link_sources = mask_template_link_sources.format(chislo=j)

    for paragraph in doc.paragraphs:
        if mask_title in paragraph.text:
            flag = True

        if flag == True:
            if paragraph.style.name == 'List Paragraph':
                sourcescount += 1
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
    # print(sourcescount, linksourcecount)


def check_pic_merged(file):
    piccount = 0
    linkpiccount = 0
    i = 1
    mask_template_pic = """Рисунок {chislo} –"""  # Задаем маску для последующей проверки по ней
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

    if piccount != linkpiccount:
        for paragraph in doc.paragraphs:    # пишем так, чтоб вывод был на поля 1 страницы
            comment = paragraph.add_comment('Кол-во иллюстраций и кол-во ссылок на иллюстрации не совпадает!')
            comment.author = 'bot'
            break
    # print(piccount, linkpiccount)


def check_table_merged(file):
    tablescount = 0
    linktablescount = 0
    i = 1
    mask_template = """Таблица {chislo} –"""  # Задаем маску для последующей проверки по ней
    mask = mask_template.format(chislo=i)  # Добавляет переменную в маску
    # Проходим по всем таблицам в документе
    for table in doc.tables:  # for table in doc.tables[1:]:
        tablescount += 1
        pass
    for paragraph in doc.paragraphs:  # Проходит по всем абзацам
        if mask in paragraph.text:  # Проверяет есть ли маска в тексте
            linktablescount += 1
            i += 1
            mask = mask_template.format(chislo=i)

    if tablescount != linktablescount:  # Если на титульном листе будут таблицы, то к ним не идут ссылки.
        for paragraph in doc.paragraphs:
            comment = paragraph.add_comment('Кол-во таблиц и кол-во ссылок на таблицы не совпадает!')
            comment.author = 'bot'
            break
    # print(tablescount, linktablescount)


def bot_comment(paragraph, TEXT: str):
    comment = paragraph.add_comment(TEXT)
    comment.author = 'bot'


def check_block(file):
    referat = False
    mask_referat = """РЕФЕРАТ"""

    soderjanie = False
    mask_soderjanie = """СОДЕРЖАНИЕ"""

    vvedenie = False
    mask_vvedenie = """ВВЕДЕНИЕ"""

    osnov_chasti = False
    mask_osnov_chasti = """ОСНОВНАЯ ЧАСТЬ"""

    zakluchenie = False
    mask_zakluchenie = """ЗАКЛЮЧЕНИЕ"""

    bibliogr_spisok = False
    mask_bibliogr_spisok = """БИБЛИОГРАФИЧЕСКИЙ СПИСОК"""

    for paragraph in doc.paragraphs:
        if paragraph.style.name == 'Heading 1':
            if mask_referat in paragraph.text:
                referat = True
                # print(paragraph.text)
            elif mask_soderjanie in paragraph.text:
                soderjanie = True
                # print(paragraph.text)
            elif mask_vvedenie in paragraph.text:
                vvedenie = True
                # print(paragraph.text)
            elif mask_osnov_chasti in paragraph.text:
                osnov_chasti = True
                # print(paragraph.text)
            elif mask_zakluchenie in paragraph.text:
                zakluchenie = True
                # print(paragraph.text)
            elif mask_bibliogr_spisok in paragraph.text:
                bibliogr_spisok = True
                # print(paragraph.text)

    # if (referat != True):
    #     for paragraph in doc.paragraphs:
    #         comment = paragraph.add_comment('Блок РЕФЕРАТ отсутствует!')
    #         comment.author = 'bot'
    #         break
    #     print("Блок РЕФЕРАТ отсутствует!")

    if (referat != True):
        for paragraph in doc.paragraphs:
            bot_comment(paragraph, "Блок РЕФЕРАТ отсутствует!")
            break
        print("Блок РЕФЕРАТ отсутствует!")

    if (soderjanie != True):
        for paragraph in doc.paragraphs:
            bot_comment(paragraph, "Блок СОДЕРЖАНИЕ отсутствует!")
            break
        print("Блок СОДЕРЖАНИЕ отсутствует!")

    if (vvedenie != True):
        for paragraph in doc.paragraphs:
            bot_comment(paragraph, "Блок ВВЕДЕНИЕ отсутствует!")
            break
        print("Блок ВВЕДЕНИЕ отсутствует!")

    if (osnov_chasti != True):
        for paragraph in doc.paragraphs:
            bot_comment(paragraph, "Блок ОСНОВНАЯ ЧАСТЬ отсутствует!")
            break
        print("Блок ОСНОВНАЯ ЧАСТЬ отсутствует!")

    if (zakluchenie != True):
        for paragraph in doc.paragraphs:
            bot_comment(paragraph, "Блок ЗАКЛЮЧЕНИЕ отсутствует!")
            break
        print("Блок ЗАКЛЮЧЕНИЕ отсутствует!")

    if (bibliogr_spisok != True):
        for paragraph in doc.paragraphs:
            bot_comment(paragraph, "Блок БИБЛИОГРАФИЧЕСКИЙ СПИСОК отсутствует!")
            break
        print("Блок БИБЛИОГРАФИЧЕСКИЙ СПИСОК отсутствует!")


# check_sources_merged(doc)
# check_table_merged(doc)
# check_pic_merged(doc)

# check_num_sources(doc)
# check_link_sources(doc)

# check_num_tables(doc)
# check_link_tables(doc)

# check_num_pic(doc)
# check_link_pic(doc)

# check_block(doc)