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

    filetype: str
    filename: str
    all_data = pd.DataFrame()

    def __init__(self, filetype: str, filename: str):
        """ выбор директории и формирование списка файлов"""
        self.filetype = ("." + filetype).lower()
        self.filename = filename.lower()
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
            return pd.read_csv(file_name, sep=";", encoding="cp1251", encoding_errors="replace")

    def tables_in_one(self):
        "таблицы в одну из списка файлов"
        for file_name in self.filesnames:
            self.all_data = pd.concat([self.all_data, self.read_one_e_file(file_name)], ignore_index=True)
            print(f"считали файл: {file_name}")

    def wrile_data(self):
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


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    pos = Access_baza()

    print("собираем файлы CSV")
    y = files_to_one("csv", "All_month")
    y.all_data[["Номер терминала"]] = y.all_data[["Номер терминала"]].astype("str")
    itog = y.all_data.merge(pos.datatable, how="left", left_on="Номер терминала", right_on="posnumber",
                            suffixes=('_alfa', '_pos'))
    y.all_data = itog
    y.wrile_data()

    exit(0)

    print("собираем файлы Excel")
    x = files_to_one("xlsx", "All_month")
    x.all_data[["Код терминала", "Наименование эмитента карты"]] = x.all_data[["Код терминала", "Наименование эмитента карты"]].astype("str")
    itog = x.all_data.merge(pos.datatable, how="left", left_on="Код терминала", right_on="posnumber", suffixes=('_alfa', '_pos'))
    itog.drop(["Торговая точка", "posnumber"], axis=1, inplace=True)
    itog.rename(columns={"Дата": "Дата и время", "Наименование эмитента карты": "Банк", "magname": "Магазин", "nomer": "Номер магазина", "Namesm": "Юрлицо"}, inplace=True)
    itog["Дата"] = itog["Дата и время"].dt.date

    print(itog.dtypes)

    x.all_data = itog
    x.wrile_data()



