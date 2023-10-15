import docx

doc = docx.Document('1.docx')
tablescount = 0

# Пройдите по всем таблицам в документе
for table in doc.tables[1:]:
    # Выполните некоторые действия с таблицей...
    tablescount += 1
    pass

linktablescount = 0
i = 1
mask_template = """Таблица {chislo} —""" #Задаем маску для последующей проверки по ней
mask = mask_template.format(chislo=i) #Добавляет переменную в маску

for paragraph in doc.paragraphs: #Проходит по всем абзацам
    text = paragraph.text #Записывает текст всего абзаца в переменную
    if mask in text: #Проверяет есть ли маска в тексте
        linktablescount += 1
        i += 1
        mask = mask_template.format(chislo=i)
    text = ''

piccount = 0
linkpiccount = 0
j = 1

#Проходит по всем иллюстрациям в документе
for shape in doc.inline_shapes:
    piccount += 1
    pass

mask_template_pic = """Рисунок {chislo} —""" #Задаем маску для последующей проверки по ней
mask_pic = mask_template_pic.format(chislo=j) #Добавляет переменную в маску

for paragraph in doc.paragraphs: #Проходит по всем абзацам
    text = paragraph.text #Записывает текст всего абзаца в переменную
    if mask_pic in text: #Проверяет есть ли маска в тексте
        linkpiccount += 1
        j += 1
        mask_pic = mask_template_pic.format(chislo=j)
    text = ''

# Вывести количество таблиц
print("Количество таблиц(все):", len(doc.tables)) #Считает все таблицы + та что на титульном
print("Количество ссылок на таблицы:", linktablescount)
print("Количество таблиц:", tablescount)
print("Количество ссылок на иллюстрации:", linkpiccount)
print("Количество иллюстраций:", piccount)

# Не работает, надо чета делать :(
def check_links():
    try:
        if linktablescount != tablescount:
            print('Кол-во таблиц и ссылок не совпадает ✕')
        elif linktablescount == tablescount:
            print('Кол-во таблиц и ссылок совпадает ✓')
        elif linkpiccount != piccount:
            print('Кол-во иллюстраций и ссылок не совпадает ✕')
        elif linkpiccount == piccount:
            print('Кол-во иллюстраций и ссылок совпадает ✓')
    except Exception as exc:
        print('Ошибка check_margin!')

#1)Кол-во табл = кол-ву ссылок (Таблица 1 - ....)
#2)Проверка "Таблца 1 -" маска
#3)Проверка ссылки на талблицу в тексте (..... в таблцие 1)

#1)Кол-во илюстрации = кол-ву ссылок (Рисунок 1 - ....)
#2)Проверка "Рисунок 1 -" маска
#3)Проверка ссылки на талблицу в тексте (..... в таблцие 1)