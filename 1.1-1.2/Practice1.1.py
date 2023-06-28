import openpyxl
import sys
#Вариант 5, задание 1
print("Практика 1.1")
#Сорок разбойников сдали экзамен по охране окружающей среды. Использовать функции
#обработки списков, чтобы найти, сколько разбойников будут охранять среду отлично, сколько
#хорошо, а сколько посредственно. Использовать функции обработки списков.
print("Задание 1")


def read_second_column(filename):
    values = []  #Список для хранения значений из второго столбца

    #Открываем файл Excel
    workbook = openpyxl.load_workbook(filename)
    #Выбираем активный лист
    sheet = workbook.active

    #Проходим по всем ячейкам второго столбца
    for row in sheet.iter_rows(min_row=1, min_col=2, max_col=2):
        for cell in row:
            cell_value = cell.value
            values.append(cell_value)

    #Закрываем файл Excel
    workbook.close()

    return values

#Указываем имя файла Excel
filename = "bandits.xlsx"
#Вызываем функцию для чтения значений из второго столбца
column_values = read_second_column(filename)
print("Отлично:", column_values.count("Отлично"))
print("Хорошо:", column_values.count("Хорошо"))
print("Посредственно:", column_values.count("Посредственно"))


print("Задание 2")
#В таблице хранятся данные о расходе электроэнергии в школе помесячно в течение года.
#Использовать функции обработки списков, чтобы узнать средний расход электроэнергии,
#минимальный и максимальный расходы, а также узнать, на сколько процентов отличаются
#минимальный и максимальный расходы от среднемесячного
def difference(avg, minmax): #функция, высчитывающая сколько в процентах minmax составляет от avg
    x = (minmax * 100) / avg
    return x

#аналогично с первым заданием
filename = "bills.xlsx"
electricity = read_second_column(filename)

electricity_avg = round(sum(electricity) / len(electricity))#среднее
electricity_min = min(electricity)#минимальное
electricity_max = max(electricity)#максимальное
#расчет на сколько процентов мин/макс отличается от среднего
min_difference = round((100 - difference(electricity_avg, electricity_min)), 1)
max_difference = round((difference(electricity_avg, electricity_max) - 100), 1)


print("Средний:", electricity_avg)
print("Минимальный:", electricity_min)
print("Максимальный:", electricity_max)
print("Среднее больше минимального на", min_difference, "процентов.")
print("Среднее меньше максимального на", max_difference, "процентов.")





