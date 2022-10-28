# Слияние файлов Excel в один. Файлы должны лежать в одной директории и иметь одинаковую структуру данных.
import os
import pandas as pd
from tkinter import filedialog
import pyodbc


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
        dirname = filedialog.askdirectory(initialdir="d:\\OneDrive\\Рабочие документы\\Эквайринг\\", title=f"Выбор кталога с файлами {filetype}").replace("/", chr(92)) # initialdir=os.getcwd()
        self.dirname = dirname
        all_dir = os.listdir(dirname)
        self.filesnames = [dirname + chr(92) + f for f in all_dir if os.path.isfile(dirname + chr(92) + f) and f[-len(self.filetype):].lower() == self.filetype.lower()]
        self.tables_in_one()


    def read_one_e_file(self, file_name):
        """ читает файл ексель (первую страницу) и возвращает датафрейм """
        if self.filetype in [".xls", ".xlsx", ".xlsm", ".xlsb", ".odf", ".ods", ".odt"]:
            return pd.read_excel(pd.ExcelFile(file_name))
        elif self.filetype == ".csv":
            return pd.read_csv(file_name, sep=";", encoding="cp1251", encoding_errors="replace", skiprows=(None if self.skp_row==0 else self.skp_row),
                               header=(None if self.head_new else "infer"), names=(self.table_header if self.head_new else None))

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
    eqv.all_data[["Номер терминала", "ID оплаты", "Референс операции", "Банк", "Валюта"]] = eqv.all_data[["Номер терминала", "ID оплаты", "Референс операции", "Банк", "Валюта"]].astype("int64").astype("str")
    eqv.all_data = eqv.all_data.drop_duplicates(subset=['Дата', 'Время', 'Карта или счет', 'Номер терминала', "Код авторизации",'Сумма транзакции'])
    eqv.wrile_data()
    full_data = pd.concat([full_data, eqv.all_data], ignore_index=True)
    print("эквайринг собрали")

    print("собираем файлы CSV системы быстрых платежей **************")
    sbp = files_to_one("csv", "All_month_sbp")
    sbp.all_data = sbp.all_data.fillna("")
    new_header = ['Дата', 'Время', 'ID СБП', 'ID QR', 'Наименование ЮЛ', 'ID ТСП', 'Карта или счет', 'Тип QR',
                  'Банк', 'Сумма транзакции', 'Тип операции', 'Номер терминала', 'Комиссия банка', 'Сумма перевода',
                  'Валюта', 'Назначение платежа', 'Имя магазина', 'ID магазина', 'ID оплаты', 'Референс операции']
    sbp.all_data.columns = new_header

    sbp.all_data["Дата"] = pd.to_datetime(sbp.all_data["Дата"], format='%d.%m.%Y')
    sbp.all_data[["Номер терминала", "Банк"]] = sbp.all_data[["Номер терминала", "Банк"]].astype("str")
    # sbp.all_data["ID оплаты"] = sbp.all_data["ID оплаты"].astype("str")
    sbp.all_data = sbp.all_data.drop_duplicates(subset=['Дата', 'Время', 'ID СБП', 'ID QR', 'Карта или счет', 'Номер терминала', 'Сумма транзакции'])
    # eqv.all_data["Дубль"] = eqv.all_data.duplicated(subset=['Дата', 'Время', 'Карта или счет', 'Номер терминала', "Код авторизации",'Сумма транзакции'])

    sbp.wrile_data()
    full_data = pd.concat([full_data, sbp.all_data], ignore_index=True)

    print("СБП собрали")

    # print("эквайринг", eqv.all_data.dtypes)
    # print("СБП",sbp.all_data.dtypes)
    print("объединение",full_data.dtypes)
    print(full_data[['Дата', 'Время']])
    # full_data.to_excel("\\".join(sbp.dirname.split("\\")[:-1]) + chr(92) + "full_data.xlsx", sheet_name="Sheet0", index=False)
    # exit(0)





    # не забыть удалить объединенный файл из каталога, если он уже есть
    print("собираем файлы Excel выгрузки из БК за месяц *************************")
    exl = files_to_one("xlsx", "All_BK_month")
    exl.all_data = exl.all_data.fillna("")
    print("проверка", exl.all_data.dtypes)
    print(exl.all_data)
    tmp_data = pd.to_datetime(exl.all_data["Дата"], format='%d/%m/%Y %H:%M:%S')
    exl.all_data["Дата"] = pd.to_datetime(tmp_data.dt.date)
    exl.all_data["Время"] = tmp_data.dt.time


    print("*****************", exl.all_data.dtypes)
    print(exl.all_data[['Дата', 'Время']])

    new_header = ["Дата", "Платежная система", "Карта или счет", "Имя магазина", "Номер терминала", "Банк",
                  "Код авторизации", "Тип операции", "Комиссия банка", "Валюта комиссии", "Сумма транзакции",
                  "Валюта", "Статус", "Время"]
    exl.all_data.columns = new_header
    exl.all_data.drop(["Валюта комиссии", "Статус"], axis= 1, inplace=True )
    new_header = ["Дата", "Время", "Платежная система", "Карта или счет", "Имя магазина", "Номер терминала", "Банк",
                  "Код авторизации", "Тип операции", "Комиссия банка", "Сумма транзакции", "Валюта"]
    exl.all_data = exl.all_data.reindex(columns=new_header)
    exl.all_data["Номер терминала"] = exl.all_data["Номер терминала"].astype("str")
    exl.all_data["Комиссия банка"] = exl.all_data["Комиссия банка"] * (-1)

    def chg_time(time_wrong: str):
        time_delt = 3
        time_wrong = time_wrong.strftime("%H:%M:%S")
        time_new = int(time_wrong[:2]) + time_delt
        if time_new >= 24:
            raise ("ошибка конвертации времени")
        return f"{time_new:02}{time_wrong[-6:]}"

    exl.all_data['Время'] = exl.all_data['Время'].apply(chg_time)

    # itog = x.all_data.merge(pos.datatable, how="left", left_on="Номер терминала", right_on="posnumber", suffixes=('_alfa', '_pos'))
    # itog.drop(["Торговая точка", "posnumber"], axis=1, inplace=True)
    # itog.rename(columns={"Дата": "Дата и время", "Наименование эмитента карты": "Банк", "magname": "Магазин", "nomer": "Номер магазина", "Namesm": "Юрлицо"}, inplace=True)
    #
    exl.wrile_data()

    def chg_operation(per: str):
        per_up = per.upper()
        if per_up == "КРЕДИТ":
            return "Покупка"
        elif per_up == "ДЕБЕТ":
            return "Возврат"
        else:
            return per

    full_data = pd.concat([full_data, exl.all_data], ignore_index=True)
    # full_data['Сумма перевода'] = full_data['Сумма перевода'].apply(lambda x: full_data["Комиссия банка"] + full_data["Сумма транзакции"] if type(x) != int or type(x) != float else x)
    full_data['Тип операции'] = full_data['Тип операции'].apply(chg_operation)
    for i in range(1, len(full_data)):
        if full_data.loc[i, 'Тип операции'] == "Возврат":
            full_data.loc[i, 'Сумма транзакции'] = full_data.loc[i, 'Сумма транзакции'] * -1




    print(full_data.dtypes)
    print(full_data[['Дата', 'Время']])

    full_data.to_excel("\\".join(sbp.dirname.split("\\")[:-1]) + chr(92) + "full_data.xlsx", sheet_name="Sheet0", index=False)



