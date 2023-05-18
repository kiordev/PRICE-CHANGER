import openpyxl
from articles_data_base import articles
from roznitsa import price_retail
from diller import price_wholesale

# Открытие таблицы Excel
wb = openpyxl.load_workbook('TableTest.xlsx')
sheet = wb.active

# Сравнение значений столбца "Артикул" и присваивание цен
start_row = 7
row = start_row
col_article = 'J'
col_retail = 'H'
col_wholesale = 'I'

while sheet[col_article + str(row)].value:
    article = sheet[col_article + str(row)].value
    if article in articles:
        index = articles.index(article)
        sheet[col_retail + str(row)].value = price_retail[index]
        sheet[col_wholesale + str(row)].value = price_wholesale[index]
    row += 1

# Сохранение обновленной таблицы
wb.save('обновленная_таблица.xlsx')