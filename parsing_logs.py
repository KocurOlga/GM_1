import pandas
from collections import defaultdict, Counter

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

#def getlist(element_list):
#    """функция создает список с товарами
#    :param element_list: список на входе, в котором
#    элемент списка = ключ
#    количество одинаковых элементов = значение"""
#    element_dict = defaultdict(int)
#    for element in element_list:
#        element_dict[element] += 1
#    return list(element_dict)


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
        
print(browser_list)
print(Counter(item_list).most_common(1))
print(Counter(men_item_list).most_common(1))
print(Counter(women_item_list).most_common(1))