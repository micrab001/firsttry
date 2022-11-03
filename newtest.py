import pandas as pd
# import os
# from tkinter import filedialog
import datetime

table_header = ["Дата", "Время", "ID оплаты", "Референс операции", "Номер терминала", "Имя магазина", "Карта или счет",
                "Банк", "Платежная система", "Сумма транзакции", "Тип операции", "Код авторизации", "Комиссия банка",
                "Сумма перевода", "Валюта"]

df = pd.read_csv("D:\\OneDrive\\Рабочие документы\\Эквайринг\\рабочая\\0182 26_10_2022_21_37_01 91846240.298.csv",
                 sep=";", encoding="cp1251", encoding_errors="replace", skiprows=6, header=None, names=table_header, dtype={'Код авторизации': 'str'})

print(df["Код авторизации"])
print(df.dtypes)


# test = datetime.datetime(2022, 9, 30)
# print(test)
# print(test + datetime.timedelta(days=1))
# print(test - datetime.timedelta(days=1))
#
# a = 3
# a += 1
# print(a)

# exit(0)
# magazin_baza = {"Галерея": ["Галерея Водолей", 14], "БУМ": ["Марьино БУМ", 132], "Дом 76А": ["Сокол", 104],
#                 "Отрадное": ["Отрадное", 59], "Семеновский": ["Семеновский", 194], "Коламбус": ["Пражская", 172],
#                 "МКАД": ["Вегас", 255], "Планерное": ["Планерная", 251], "Витте Молл": ["Бутово", 310],
#                 "Речной": ["Речной", 247], "Зелёный": ["Новогиреево", 301], "Калейдоскоп": ["Сходненская", 343],
#                 "Дубравная": ["Митино", 351], "Кунцево": ["Кунцево", 422], "AVENUE": ["Авеню Ю-З", 536],
#                 "Каширская": ["Каширский", 529]}
#
#
# def add_magazin(adress: str):
#     global magazin_baza
#     for mag in magazin_baza:
#         if mag in adress:
#             return (magazin_baza[mag][0], magazin_baza[mag][1])
#     return("не найден", "не найден")
#
#
# addrr = "Город Москва Город, ул Венёвская, Дом 6, ТЦ Витте Молл, Пом. 1, Ком. 24"
# print(add_magazin(addrr))

# dirname = "D:\\OneDrive\\Рабочие документы\\Эквайринг\\2022_09 Альфа\\2022_09 Альфа эквайринг"
# all_dir = os.listdir(dirname)
# if "all_bk_month.xlsx" in  all_dir:
#     os.remove(dirname + chr(92) + "all_bk_month.xlsx")
#     print("deleted")
#     all_dir.remove("all_bk_month.xlsx")

# filename = filedialog.askopenfilename(initialdir="d:\\OneDrive\\Рабочие документы\\Эквайринг\\рабочая")
# bank_operation = pd.read_excel(pd.ExcelFile(filename), "Вся выписка")
# print(filename)
# print(len(bank_operation))
# list_inn = list(map(str,bank_operation["ПолучательИНН"].unique()))
# print(bank_operation.dtypes)
# print(list_inn)
#
# bank_operation = bank_operation[bank_operation['НазначениеПлатежа'].str.contains('|'.join(list_inn))]
#
# print(len(bank_operation))
# print(bank_operation['НазначениеПлатежа'])
# full_data = pd.DataFrame({'Сумма транзакции': [1894, 511, 11, 1493, 36],
# 'Тип операции': ["Покупка", "Возврат", "Покупка", "Возврат", "Покупка"]})
#
# mag_komiss = pd.pivot_table(full_data, index=['Тип операции'], values=['Сумма транзакции'],
#                                 aggfunc=sum)
#
#
#
# print(mag_komiss["Сумма транзакции"])
# print(mag_komiss * -1)

# a = full_data["Тип операции"].unique()
# b = full_data["Сумма транзакции"].unique()
# for el in b:
#     if el in a:
#         print("совпало")

#
# b = list(full_data["Тип операции"].unique())
# print(type(b))
# print(list(b))
# print(type(list(b)))
# print(a- b)
# print(len(a-b))


# for i in range(1, len(full_data)):
#     if full_data.loc[i, 'Тип операции' ] == "Возврат":
#         full_data.loc[i, 'Сумма транзакции' ] = full_data.loc[i, 'Сумма транзакции' ] * -1

# print(full_data)