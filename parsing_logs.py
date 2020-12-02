import pandas
from collections import defaultdict, Counter
from openpyxl import load_workbook

def getvalue(object_dict, item):
    """функция, в которой выводится значение заданного ключа:
    :param object_dict: словарь, который смотрим
    :param item: заголовок столбца, значение которого нужно получить"""
    value = None
    for k, v in object_dict.items():
        if k == item:
            value = v
            break
    return value


# читаем файл эксель построчно с указанием вкладки
logs_data = pandas.read_excel('logs1.xlsx', sheet_name='log')
# для каждой строки получаем словарь со связкой "название столбца - значение"
logs_dict = logs_data.to_dict(orient='records')

# создаем пустые списки словарей для отслеживаемых параметров
browser_list = [] #список браузеров
item_list = [] #список товаров общий
men_item_list = [] #список товаров у мужчин
women_item_list = [] #список товаров у женщин

#разбиваем строки с подсчетом нужных нам параметров
for record in logs_dict:
    # вытаскиваем товары
    browser_list.append(getvalue(record, 'Браузер'))
    item_list.extend(getvalue(record, 'Купленные товары').split(','))
    if getvalue(record, 'Пол') == 'м':
        men_item_list.extend(getvalue(record, 'Купленные товары').split(','))
    else:
        women_item_list.extend(getvalue(record, 'Купленные товары').split(','))

#записываем полученные данные в отчет
wb = load_workbook(filename='report.xlsx')
sheet = wb['Лист1']
#популярные товары среди мужчин и женщин
sheet['B31'] = Counter(men_item_list).most_common(1)[0][0]
sheet['B32'] = Counter(women_item_list).most_common(1)[0][0]
#непопулярные товары среди мужчин и женщин
len_men_counter = len(men_item_list)
sheet['B33'] = Counter(men_item_list).most_common()[:-(len_men_counter + 1): -1][0][0]
len_women_counter = len(men_item_list)
sheet['B34'] = Counter(women_item_list).most_common()[:-(len_women_counter + 1): -1][0][0]
wb.save('report.xlsx')
