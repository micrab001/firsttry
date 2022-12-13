# Слияние файлов Excel в один. Файлы должны лежать в одной директории и иметь одинаковую структуру данных.
import os
import pandas as pd
from tkinter import filedialog
import pyodbc
import datetime
import sqlite3
import requests

datefrom = datetime.datetime(2022, 11, 1)
dateto = datetime.datetime(2022, 11, 30)  # берется диапазон включая крайние даты

class files_to_one():
    """ Открывает по очереди файлы из указанной директории, соединяет таблицы вместе из этих файлов и записывает
    результирующую таблицу в один. Принимает параметры
    filetype: str - типы файлов (расширение) без точки
    filename: str - результирующий файл"""

    # filetype: str
    # filename: str
    all_data = pd.DataFrame()
    table_header = ["Дата", "Время", "ID оплаты", "Референс операции", "Номер терминала", "Имя магазина", "Карта или счет",
                    "Банк", "Платежная система", "Сумма транзакции", "Тип операции", "Код авторизации", "Комиссия банка",
                    "Сумма перевода", "Валюта"]


    def __init__(self, filetype: str, filename: str, skp_row: int = 0, head_new: bool = False):
        """ выбор директории и формирование списка файлов"""
        self.filetype = ("." + filetype).lower()
        self.filename = filename.lower()
        self.skp_row = skp_row
        self.head_new = head_new
        dirname = filedialog.askdirectory(initialdir="d:\\OneDrive\\Рабочие документы\\Эквайринг\\Альфа", title=f"Выбор кталога с файлами {filetype}").replace("/", chr(92)) # initialdir=os.getcwd()
        self.dirname = dirname
        all_dir = os.listdir(dirname)
        filename = self.filename + ".xlsx"
        if self.filetype == ".xlsx" and filename in all_dir:
            os.remove(dirname + chr(92) + filename)
            print(f"file {filename} deleted")
            all_dir.remove(filename)
        self.filesnames = [dirname + chr(92) + f for f in all_dir if os.path.isfile(dirname + chr(92) + f) and f[-len(self.filetype):].lower() == self.filetype.lower()]
        self.tables_in_one()


    def read_one_e_file(self, file_name):
        """ читает файл ексель (первую страницу) и возвращает датафрейм """
        if self.filetype in [".xls", ".xlsx", ".xlsm", ".xlsb", ".odf", ".ods", ".odt"]:
            return pd.read_excel(pd.ExcelFile(file_name))
        elif self.filetype == ".csv":
            return pd.read_csv(file_name, sep=";", encoding="cp1251", encoding_errors="replace", skiprows=(None if self.skp_row==0 else self.skp_row),
                               header=(None if self.head_new else "infer"), names=(self.table_header if self.head_new else None),
                               dtype=({'Код авторизации': 'str'} if self.head_new else None ))

    def tables_in_one(self):
        """таблицы в одну из списка файлов"""
        for file_name in self.filesnames:
            self.all_data = pd.concat([self.all_data, self.read_one_e_file(file_name)], ignore_index=True)
            print(f"считали файл: {file_name}")

    def wrile_data(self):
        """записывает собранную таблицу в один файл Excel"""
        self.all_data.to_excel(self.dirname + chr(92) + self.filename + ".xlsx", sheet_name="Sheet0", index=False)

class Access_baza():
    def __init__(self):
        """ получение из файла access табличку по терминалам и магазинам"""
        fn = "D:\\Работа\\baza\\kosmbase.mdb"
        conn_str = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" + fn + ";"
        cnxn = pyodbc.connect(conn_str)
        crsr = cnxn.cursor()
        sql = """SELECT terminals.posnumber, Magazin.magname, Magazin.nomer, Kompanii.Namesm, Kompanii.INN
                 FROM Kompanii INNER JOIN (Magazin INNER JOIN terminals ON Magazin.magkey = terminals.Магазин) ON Kompanii.orgkey = Magazin.magkomp
                 ORDER BY Magazin.magname;"""
        crsr.execute(sql)
        self.datatable = pd.DataFrame(list(map(list, crsr.fetchall())), columns=[name[0] for name in crsr.description])
        crsr.close()
        cnxn.close()

# def vopros(info: str = " "):
#     process = input(info + " 1 - да, 2 пропустить :")
#     if process == "1":
#         return True
#     else:
#         return False


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    pos = Access_baza()
    full_data = pd.DataFrame()

    print("собираем файлы CSV эквайринг ***************")
    eqv = files_to_one("csv", "All_month_eqv", 6, True)
    eqv.all_data = eqv.all_data[eqv.all_data["Дата"].notnull()]
    eqv.all_data = eqv.all_data.fillna("")
    eqv.all_data["Дата"] = pd.to_datetime(eqv.all_data["Дата"], format='%d.%m.%Y')
    eqv.all_data = eqv.all_data[(eqv.all_data["Дата"] >= datefrom) & (eqv.all_data["Дата"] <= dateto)]
    eqv.all_data[["Номер терминала", "ID оплаты", "Референс операции", "Банк", "Валюта"]] = eqv.all_data[["Номер терминала", "ID оплаты", "Референс операции", "Банк", "Валюта"]].astype("int64").astype("str")
    eqv.all_data = eqv.all_data.drop_duplicates(subset=['Дата', 'Время', 'Карта или счет', 'Номер терминала', "Код авторизации",'Сумма транзакции'])
    eqv.all_data["Система"] = "эквайринг"
    print("**************** эквайринг записываем общий файл")
    eqv.wrile_data()
    full_data = pd.concat([full_data, eqv.all_data], ignore_index=True)
    print("**************** эквайринг собрали")

    print("собираем файлы CSV системы быстрых платежей **************")
    sbp = files_to_one("csv", "All_month_sbp")
    sbp.all_data = sbp.all_data.fillna("")
    new_header = ['Дата', 'Время', 'ID СБП', 'ID QR', 'Наименование ЮЛ', 'ID ТСП', 'Карта или счет', 'Тип QR',
                  'Банк', 'Сумма транзакции', 'Тип операции', 'Номер терминала', 'Комиссия банка', 'Сумма перевода',
                  'Валюта', 'Назначение платежа', 'Имя магазина', 'ID магазина', 'ID оплаты', 'Референс операции']
    sbp.all_data.columns = new_header
    sbp.all_data["Дата"] = pd.to_datetime(sbp.all_data["Дата"], format='%d.%m.%Y')
    sbp.all_data = sbp.all_data[(sbp.all_data["Дата"] >= datefrom) & (sbp.all_data["Дата"] <= dateto)]
    sbp.all_data[["Номер терминала", "Банк", "ID оплаты"]] = sbp.all_data[["Номер терминала", "Банк", "ID оплаты"]].astype("str")
    sbp.all_data["Система"] = "СБП"
    sbp.all_data = sbp.all_data.drop_duplicates(subset=['Дата', 'Время', 'ID СБП', 'ID QR', 'Карта или счет', 'Номер терминала', 'Сумма транзакции'])
    print("********************************* СБП записываем общий файл")
    sbp.wrile_data()
    full_data = pd.concat([full_data, sbp.all_data], ignore_index=True)
    print("********************************* СБП собрали")

    # не забыть удалить объединенный файл из каталога, если он уже есть

    print("собираем файлы Excel выгрузки из БК за месяц *************************")
    exl = files_to_one("xlsx", "All_BK_month")
    exl.all_data = exl.all_data.fillna("")
    tmp_data = pd.to_datetime(exl.all_data["Дата"], format='%d/%m/%Y %H:%M:%S')
    exl.all_data["Дата"] = pd.to_datetime(tmp_data.dt.date)
    exl.all_data["Время"] = tmp_data.dt.time
    exl.all_data = exl.all_data[(exl.all_data["Дата"] >= datefrom) & (exl.all_data["Дата"] <= dateto)]
    new_header = ["Дата", "Платежная система", "Карта или счет", "Имя магазина", "Номер терминала", "Банк",
                  "Код авторизации", "Тип операции", "Комиссия банка", "Валюта комиссии", "Сумма транзакции",
                  "Валюта", "Статус", "Время"]
    exl.all_data.columns = new_header
    exl.all_data.drop(["Валюта комиссии", "Статус"], axis= 1, inplace=True )
    new_header = ["Дата", "Время", "Платежная система", "Карта или счет", "Имя магазина", "Номер терминала", "Банк",
                  "Код авторизации", "Тип операции", "Комиссия банка", "Сумма транзакции", "Валюта"]
    exl.all_data = exl.all_data.reindex(columns=new_header)
    exl.all_data["Система"] = "эквайринг"
    exl.all_data[["Номер терминала", "Код авторизации"]] = exl.all_data[["Номер терминала", "Код авторизации"]].astype("str")
    exl.all_data["Комиссия банка"] = exl.all_data["Комиссия банка"] * (-1)

    def chg_time(time_wrong: str):
        time_delt = 3
        time_wrong = time_wrong.strftime("%H:%M:%S")
        time_new = int(time_wrong[:2]) + time_delt
        if time_new >= 24:
            raise ("ошибка конвертации времени")
        return f"{time_new:02}{time_wrong[-6:]}"

    exl.all_data['Время'] = exl.all_data['Время'].apply(chg_time)
    print("**************************************** excel записываем общий файл")
    exl.wrile_data()
    print("**************************************** excel собрали")

    def chg_operation(per: str):
        per_up = per.upper()
        if per_up == "КРЕДИТ":
            return "Покупка"
        elif per_up == "ДЕБЕТ":
            return "Возврат"
        else:
            return per

    print("Обрабатываем все операции ********************************")
    full_data = pd.concat([full_data, exl.all_data], ignore_index=True)
    # проверка списка терминалов, чтобы все из отчета были в таблице из базы
    term_baza = pos.datatable["posnumber"].unique()  # получаем список терминалов из базы
    for term_otchet in full_data["Номер терминала"].unique():  # ищем терминалы из отчета в списке терминалов в базе
        if term_otchet not in term_baza:
            print(f"************************ незарегистрированный терминал №{term_otchet} . Требуется сначала внести его в базу")
            exit("найден неопознанный терминал")
    full_data['Тип операции'] = full_data['Тип операции'].apply(chg_operation)

    db = sqlite3.connect("c:\\Vrem\\Python10\\bd_bin_code.db")
    cur = db.cursor()


    def convert_bin(val):
        if type(val) != str:
            return "Другой"
        bin_code = val[0:6]
        sql_str = f"SELECT * FROM bincode WHERE BIN = '{bin_code}';"
        cur.execute(sql_str)
        rezult = cur.fetchall()
        col_name = [el[0] for el in cur.description]
        if len(rezult) == 1:
            rez = dict(zip(col_name, rezult[0]))
            if rez['Банк-эмитент'] != "":
                return rez['Банк-эмитент']
            else:
                return "Другой"
        else:
            return "Другой"


    def convert_bic_code(val):
        if type(val) != str:
            return "Другой"
        bic_code = val
        sql_str = f"SELECT * FROM bic_code WHERE BIC LIKE '%{bic_code}';"
        cur.execute(sql_str)
        rezult = cur.fetchall()
        col_name = [el[0] for el in cur.description]
        if len(rezult) == 1:
            rez = dict(zip(col_name, rezult[0]))
            if rez['Name_org'] != "":
                return rez['Name_org']
            else:
                return "Другой"
        else:
            return "Другой"


    full_data.insert(full_data.columns.get_loc('Банк') + 1, 'Наименование банка', "")
    count = 1
    flag = False
    rez = ""
    new_bin = []
    bin_not_found = []
    chk_calc = []
    for i in range(0, len(full_data)):
        # Печать счетчика
        procent = int(i/len(full_data)*100)
        if procent >= count:
            if flag:
                print("\b" * len(rez), end="", flush=True)
            rez = f"Обрабатываем операции: {procent}%"
            print(rez, end ="")
            count += 1
            flag = True

        if full_data.loc[i, 'Тип операции'] == "Возврат" and full_data.loc[i, 'Сумма транзакции'] > 0:
            full_data.loc[i, 'Сумма транзакции'] = full_data.loc[i, 'Сумма транзакции'] * -1
            if full_data.loc[i, "Система"] == "СБП":
                full_data.loc[i, "Сумма перевода"] = full_data.loc[i, "Сумма перевода"] * -1
        if (pd.isna(full_data.loc[i, "Код авторизации"]) or full_data.loc[i, "Код авторизации"] == "") and full_data.loc[i, "Система"] == "эквайринг":
            full_data.loc[i, "Код авторизации"] = "000000"
            if full_data.loc[i, "Комиссия банка"] > 0:
                full_data.loc[i, 'Сумма транзакции'] = full_data.loc[i, 'Сумма транзакции'] * -1
        if pd.isna(full_data.loc[i, "Сумма перевода"]):
            full_data.loc[i, "Сумма перевода"] = full_data.loc[i, 'Сумма транзакции'] + full_data.loc[i, 'Комиссия банка']
        if len(full_data.loc[i, "Тип операции"]) == "Покупка":
            if full_data.loc[i, "Система"] == "СБП" and round(full_data.loc[i, "Сумма транзакции"] * 0.007, 2) != abs(
                    full_data.loc[i, "Комиссия банка"]):
                chk_calc.append(
                    f'{full_data.loc[i, "Дата"]}, {full_data.loc[i, "Время"]}, {full_data.loc[i, "Номер терминала"]}, {full_data.loc[i, "Сумма транзакции"]}')
            elif round(full_data.loc[i, "Сумма транзакции"] * 0.012, 2) != abs(full_data.loc[i, "Комиссия банка"]):
                chk_calc.append(
                    f'{full_data.loc[i, "Дата"]}, {full_data.loc[i, "Время"]}, {full_data.loc[i, "Номер терминала"]}, {full_data.loc[i, "Сумма транзакции"]}')
        if len(full_data.loc[i, "Банк"]) > 7 and full_data.loc[i, "Система"] == "СБП":
            full_data.loc[i, "Наименование банка"] = convert_bic_code(full_data.loc[i, "Банк"])
        else:
            full_data.loc[i, "Наименование банка"] = convert_bin(full_data.loc[i, "Карта или счет"])
            if full_data.loc[i, "Наименование банка"].upper() == "ДРУГОЙ":
                if full_data.loc[i, "Банк"] == "0":
                    sql_str = f'INSERT OR REPLACE INTO bincode ("BIN", "Платежная система", "Страна", "Банк-эмитент", "Адрес сайта банка")' \
                              f' VALUES ("{full_data.loc[i, "Карта или счет"][0:6]}", "{full_data.loc[i, "Платежная система"].split(" ")[0]}",' \
                              f' "{"Россия"}", "{"Alfa-Bank, Альфа-Банк"}", "{"alfabank.ru"}");'
                    cur.execute(sql_str)
                    db.commit()
                    new_bin.append(f'добавили бин {full_data.loc[i, "Карта или счет"][0:6]} Альфабанка')
                    full_data.loc[i, "Наименование банка"] = convert_bin(full_data.loc[i, "Карта или счет"])
                else:
                    if full_data.loc[i, 'Карта или счет'][0:6] not in bin_not_found:
                        url = "https://bin-ip-checker.p.rapidapi.com/"
                        querystring = {"bin": f"{full_data.loc[i, 'Карта или счет'][0:6]}"}
                        payload = {"bin": f"{full_data.loc[i, 'Карта или счет'][0:6]}"}
                        headers = {"content-type": "application/json",
                                   "X-RapidAPI-Key": "2ee170771emshe310361e8baaa6dp1b01dejsn7b3d6d68abb8",
                                   "X-RapidAPI-Host": "bin-ip-checker.p.rapidapi.com"}
                        response = requests.request("POST", url, json=payload, headers=headers, params=querystring)
                        if response.status_code != 404:
                            answ_text = response.json()
                            if answ_text["success"] and answ_text["code"] == 200:
                                if answ_text["BIN"]["valid"] and answ_text["BIN"]["issuer"]["name"] != "":
                                    sql_str = f'INSERT OR REPLACE INTO bincode ("BIN", "Платежная система", "Страна", "Банк-эмитент", "Тип карты", "Категория карты", "Адрес сайта банка")' \
                                              f' VALUES ("{answ_text["BIN"]["number"]}", "{answ_text["BIN"]["brand"]}",' \
                                              f' "{answ_text["BIN"]["country"]["country"]}", "{answ_text["BIN"]["issuer"]["name"]}",' \
                                              f' "{answ_text["BIN"]["type"]}", "{answ_text["BIN"]["level"]}", "{answ_text["BIN"]["issuer"]["website"]}");'
                                    cur.execute(sql_str)
                                    db.commit()
                                    new_bin.append(
                                        f'добавили бин {full_data.loc[i, "Карта или счет"][0:6]} , банк {answ_text["BIN"]["issuer"]["name"]} система {answ_text["BIN"]["brand"]} поиск {len(new_bin)+1}')
                                    full_data.loc[i, "Наименование банка"] = convert_bin(full_data.loc[i, "Карта или счет"])
                                else:
                                    new_bin.append(f'бин {full_data.loc[i, "Карта или счет"][0:6]} не найден, поиск {len(new_bin)+1}')
                                    bin_not_found.append(full_data.loc[i, "Карта или счет"][0:6])
                        else:
                            bin_not_found.append(full_data.loc[i, "Карта или счет"][0:6])

        # тут можно еще разных обработок наделать, например банка

    print("\b" * len(rez), end="", flush=True)
    print("Обработано 100% операций")
    print("произведен поиск следующих бинов:")
    print(*new_bin, sep="\n")
    print("ошибки начисления комиссии:")
    if len(chk_calc) == 0:
        print("не найдены")
    else:
        print(*chk_calc, sep="\n")

    full_data = full_data.drop_duplicates(subset=['Дата', 'Время', 'Карта или счет', 'Номер терминала', "Код авторизации",'Сумма транзакции'])
    full_data = full_data.merge(pos.datatable, how="left", left_on="Номер терминала", right_on="posnumber", suffixes=('_alfa', '_pos'))
    full_data.drop(["Имя магазина", "posnumber", "Наименование ЮЛ"], axis=1, inplace=True)
    full_data.rename(columns={"magname": "Магазин", "nomer": "Номер магазина", "Namesm": "Юрлицо"}, inplace=True)
    full_data.sort_values(by=["Юрлицо", "Номер магазина", "Номер терминала", 'Дата', 'Время'], inplace=True)
    full_data["Номер магазина"] = full_data["Номер магазина"].astype("str")

    def okrugl(a):
        return round(a,2)

    # загружаем данные банка по операциям
    print("Загрузка данных из банка ********************************")
    filename = filedialog.askopenfilename(initialdir="d:\\OneDrive\\Рабочие документы\\Выписки Альфа", title="Выбрать файл с выписками из банка")
    bank_operation = pd.read_excel(pd.ExcelFile(filename), "Вся выписка")
    bank_operation = bank_operation[["СекцияДокумент", "Номер", "Дата", "Сумма", "Плательщик", "Получатель", "НазначениеПлатежа"]]
    list_inn = list(map(str, full_data["INN"].unique()))
    bank_operation = bank_operation[bank_operation['НазначениеПлатежа'].str.contains('|'.join(list_inn))]
    bank_operation = bank_operation[(bank_operation["Дата"] >= datefrom) & (bank_operation["Дата"] <= (dateto + datetime.timedelta(days=1)))]
    bank_operation["Найдено"] = False

    # собираем эквайринг по дням
    print("Собираем и проверяем суммы по дням  ********************************")
    data_by_day = full_data[["INN", "Юрлицо", "Дата", "Система", 'Сумма транзакции', 'Комиссия банка',
                              'Сумма перевода']].groupby(["INN", "Юрлицо", "Дата", "Система"], as_index=False).sum()
    data_by_day["Найдено"] = False
    for i in range(0, len(data_by_day)):
        if data_by_day.loc[i, 'Система'] == "эквайринг":
            find_operation = bank_operation[(bank_operation["Дата"] == (data_by_day.loc[i, "Дата"] + datetime.timedelta(days=1))) &
                                            (bank_operation["Сумма"] == round(data_by_day.loc[i, "Сумма перевода"],2)) &
                                            (bank_operation['НазначениеПлатежа'].str.contains(data_by_day.loc[i, "INN"])) &
                                            (bank_operation["Найдено"] == False)]
            if len(find_operation) == 1:
                data_by_day.loc[i, 'Найдено'] = True
                bank_operation.loc[find_operation.index[0], "Найдено"] = True

                # bank_operation.loc[(bank_operation["Дата"] == (data_by_day.loc[i, "Дата"] + datetime.timedelta(days=1))) &
                #                (bank_operation["Сумма"] == round(data_by_day.loc[i, "Сумма перевода"],2)) &
                #                (bank_operation['НазначениеПлатежа'].str.contains(data_by_day.loc[i, "INN"])) &
                #                (bank_operation["Найдено"] == False), "Найдено"] = True

        elif data_by_day.loc[i, 'Система'] == "СБП":
            find_operation = full_data[(full_data["Дата"] == data_by_day.loc[i, "Дата"]) &
                                       (full_data["INN"] == data_by_day.loc[i, "INN"]) &
                                       (full_data['Система'] == "СБП")]
            if len(find_operation) > 0:
                sum_by_day = 0
                for j in range(0, len(find_operation)):
                    bank_operation.loc[bank_operation['НазначениеПлатежа'].str.contains(find_operation.iloc[j]["Референс операции"]), "Найдено"] = True
                    sum_by_day += bank_operation.loc[bank_operation['НазначениеПлатежа'].str.contains(find_operation.iloc[j]["Референс операции"]), "Сумма"].sum()
                sum_by_day = round(sum_by_day, 2)
                if sum_by_day == round(data_by_day.loc[i, 'Сумма транзакции'] - data_by_day.loc[i, 'Комиссия банка'], 2):
                    data_by_day.loc[i, 'Найдено'] = True

    # собираем комиссию для Филиппа
    print("Собираем комиссию Филиппу  ********************************")
    mag_komiss = pd.pivot_table(full_data, index=['Магазин'], columns=["Система"], values=['Комиссия банка'],
                                aggfunc=sum)  # margins=True
    mag_komiss = mag_komiss * -1

    print("Собираем платежи по банкам  ********************************")
    banks_data = pd.pivot_table(full_data, index=['Наименование банка'], values=['Сумма транзакции'], aggfunc=[sum, len])  # margins=True
    banks_data = banks_data.sort_values(by=[('sum', 'Сумма транзакции')], ascending=False)
    banks_data["% по сумме"] = round(banks_data["sum"] / banks_data["sum"].sum(), 4)
    banks_data["% по количеству"] = round(banks_data["len"] / banks_data["len"].sum(), 4)
    banks_data.loc['Всего'] = banks_data.sum()
    banks_data.rename(columns={'sum': 'Сумма операций', 'len': 'Количество операций'}, inplace=True)
    cols = list(banks_data.columns.values)
    cols = [cols[0], cols[2], cols[1], cols[3]]
    banks_data = banks_data[cols]
    new_header = [el[0] for el in cols]


    # записываем результат в файл
    print("Запись результирующего файла  ********************************")
    # writer = pd.ExcelWriter("\\".join(sbp.dirname.split("\\")[:-1]) + chr(92) + "full_data.xlsx", engine='xlsxwriter')
    # full_data.to_excel(writer, sheet_name='Sheet0', index=False)
    # data_by_day.to_excel(writer, sheet_name='По дням', index=False)
    # mag_komiss.to_excel(writer, sheet_name='Для Филиппа')
    # bank_operation.to_excel(writer, sheet_name="Банк", index=False)
    # # df3.to_excel(writer, sheet_name='Sheetc')
    # writer.save()

    writer = pd.ExcelWriter("\\".join(sbp.dirname.split("\\")[:-1]) + chr(92) + "full_data.xlsx", engine='xlsxwriter')
    with writer as file_name:
        full_data.to_excel(file_name, sheet_name="Sheet0", index=False)
        # banks_data.to_excel(file_name, sheet_name="banks")
        # Convert the dataframe to an XlsxWriter Excel object.
        data_by_day.to_excel(file_name, sheet_name='По дням проверка', index=False)
        mag_komiss.to_excel(file_name, sheet_name='Для Филиппа')
        bank_operation.to_excel(file_name, sheet_name="Банк проверка", index=False)
        banks_data.to_excel(file_name, sheet_name="banks", header=False)
        # Get the xlsxwriter objects from the dataframe writer object.
        workbook = writer.book
        worksheet = writer.sheets["banks"]
        # Add some cell formats.
        format1 = workbook.add_format({'num_format': '#,##0'})
        format2 = workbook.add_format({'num_format': '0%'})
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1})
        worksheet.set_row(0, 30, header_format)
        worksheet.write_row('B1', new_header)
        # Set the column width and format.
        worksheet.set_column('B:B', 15, format1)
        worksheet.set_column('C:C', 10, format2)
        worksheet.set_column('D:D', 15, format1)
        worksheet.set_column('E:E', 10, format2)




    db.close()

    # full_data.to_excel("\\".join(sbp.dirname.split("\\")[:-1]) + chr(92) + "full_data.xlsx", sheet_name="Sheet0", index=False)


