# Слияние файлов Excel в один. Файлы должны лежать в одной директории и иметь одинаковую структуру данных.
import os
import pandas as pd
from tkinter import filedialog

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
        self.filetype = "." + filetype
        self.filename = filename
        dirname = filedialog.askdirectory(initialdir="d:\\OneDrive\\Рабочие документы\\Эквайринг\\").replace("/", chr(92)) # initialdir=os.getcwd()
        self.dirname = dirname
        all_dir = os.listdir(dirname)
        self.filesnames = [dirname + chr(92) + f for f in all_dir if os.path.isfile(dirname + chr(92) + f) and f[-len(self.filetype):].lower() == self.filetype.lower()]
        self.tables_in_one()

    @staticmethod
    def read_one_e_file(file_name):
        """ читает файл ексель (первую страницу) и возвращает датафрейм """
        return pd.read_excel(pd.ExcelFile(file_name))

    def tables_in_one(self):
        "таблицы в одну из списка файлов"
        for file_name in self.filesnames:
            self.all_data = pd.concat([self.all_data, self.read_one_e_file(file_name)], ignore_index=True)
            print(f"считали файл: {file_name}")

    def wrile_data(self):
        self.all_data.to_excel(self.dirname + chr(92) + self.filename + ".xlsx", sheet_name="Sheet0", index=False)



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    x = files_to_one("xlsx", "All_month")
    print("stop")
    x.wrile_data()



