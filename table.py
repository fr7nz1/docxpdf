import docx

doc = docx.Document('1.docx')
# Пройдите по всем таблицам в документе
for table in doc.tables[1:]:
    # Выполните некоторые действия с таблицей...
    pass

    # Вывести количество таблиц
print("Количество таблиц:", len(doc.tables))
