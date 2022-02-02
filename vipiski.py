import pandas as pd
import json
from tkinter import filedialog
import os
import datetime

# выбор нужной директории и создание списка файлов в нем с расширением тхт
dirname = filedialog.askdirectory(initialdir=os.getcwd()).replace("/", chr(92))
all_dir = os.listdir(dirname)
filesnames = [dirname+chr(92)+f for f in all_dir if os.path.isfile(dirname+chr(92)+f) and ".txt" in f]

# задание некоторых начальных переменных
all_data = ""
filename_xls = dirname+chr(92)+"all_data.xlsx"
strt_time = datetime.date(1900, 1, 1)  # объект для конвертации даты


# конвертация даты из строки в число в формат для Excel
def str_to_data(ds):  # предполагается, что формат даты д-м-г если ошибка то берет г-м-d
    global strt_time
    if isinstance(ds, int):
        return ds
    else:
        if ds == "":
            return ""
        d = ""
        dd = []
        for i in ds:
            if i.isdigit():
                d += i
            else:
                dd.append(int(d))
                d = ""
        dd.append(int(d))
        try:
            return (datetime.date(dd[2] if dd[2] > 100 else dd[2] + 2000, dd[1], dd[0]) - strt_time).days + 2
        except ValueError:
            return (datetime.date(dd[0] if dd[0] > 100 else dd[0] + 2000, dd[1], dd[2]) - strt_time).days + 2


# конвертация даты и времени в дату (объект datetime)
def convert_data(val):
    return val.date()


# начало программы считываение выписок из тхт файлов в формате выгрузки для 1С и создание файла json
for filename in filesnames:
    with open(filename, "r") as data_file:
        flag = False
        for line in data_file:
            line = line.replace("\n", "").replace('"', "'").replace("\\", "/")
            if "СекцияДокумент=" in line:  # "Платежное поручение", "СекцияДокумент=Банковский ордер",
                # "СекцияДокумент=Платежное требование", "СекцияДокумент=Мемориальный ордер"):
                all_data += "{"
                flag = True
            if "КонецДокумента" in line:
                all_data += "}, \n"
                flag = False
            else:
                if flag:
                    tmp_list = line.split("=")
                    tmp_list[1] = tmp_list[1].upper()
                    if "ДАТА" in tmp_list[0].upper():
                        if tmp_list[1] != "":
                            tmp_list[1] = str_to_data(tmp_list[1])
                    if "Сумма" == tmp_list[0] or "Дата" in tmp_list[0] and tmp_list[1] != "":
                        all_data += f'"{tmp_list[0]}":{tmp_list[1]}, '
                    else:
                        all_data += f'"{tmp_list[0]}":"{tmp_list[1]}", '
                    if tmp_list[0] == "НазначениеПлатежа" and "ЗАЧИСЛЕНИЕ СРЕДСТВ ПО ОПЕРАЦИЯМ" in tmp_list[1] and "МЕРЧАНТ" in tmp_list[1]:
                        tmp_lst = tmp_list[1].split()
                        all_data += f'"Номер мерчанта":"{tmp_lst[tmp_lst.index("МЕРЧАНТ")+1].strip("№.")}", "Дата операции магазин":{str_to_data(tmp_lst[tmp_lst.index("РЕЕСТРА")+1].rstrip("."))}, ' \
                                    f'"Сумма комиссии":{tmp_lst[tmp_lst.index("КОМИССИЯ")+1].rstrip(".").replace(",", "")}, "Возврат1":{tmp_lst[tmp_lst.index("ПОКУПКИ")+1].split("/")[0].replace(",", "")}, ' \
                                    f'"Возврат2":{tmp_lst[tmp_lst.index("ПОКУПКИ")+1].split("/")[1].rstrip(".").replace(",", "")}, '

all_data = all_data.replace(", }", "}")
all_data = all_data[0: all_data.rfind(",")]
all_data = "[" + all_data + "]"

with open("test.json", "w") as tst_file:
    tst_file.write(all_data)

# начало работы с данными
# считывание выписок
df = pd.DataFrame(json.loads(all_data))
# создание выборки данных из выписок по эквайрингу
svod = df[["Дата","Сумма","ДатаПоступило","Получатель","НазначениеПлатежа","Номер мерчанта",
           "Дата операции магазин","Сумма комиссии","Возврат1","Возврат2"]]
svod = svod[svod["Номер мерчанта"].notnull()]
svod = svod.sort_values(by=["Номер мерчанта","Дата"])
svod["Получено"] = svod["Сумма"] + svod["Сумма комиссии"]

# считываение данных по эквайрингу от Сбера
filename = filedialog.askopenfilename(initialdir= os.getcwd())
sber_df = pd.read_excel(pd.ExcelFile(filename), "Sheet0")
sber_df["Дата зачисления"] = sber_df["Дата выгрузки в АБС"].apply(convert_data)
# Номер мерчанта, Дата операции магазин, Дата выгрузки в АБС, Сумма операции, Сумма комиссии, Сумма расчета
sber_df["Номер мерчанта"] = sber_df["Номер мерчанта"].astype("str")
# получение выборки данных из эквайринга Сбера
svod_sber_df = sber_df[["Номер мерчанта", "Дата зачисления", "Сумма операции", "Сумма комиссии", "Сумма расчета"]]
svod_sber_df = svod_sber_df.groupby(["Номер мерчанта", "Дата зачисления"], as_index=False).sum() #[["Сумма операции"]]
svod_sber_df["Дата зачисления"] = svod_sber_df["Дата зачисления"].astype("str").apply(str_to_data)
# слияние данных сбера и банка
itog = svod_sber_df.merge(svod, how = "left", left_on= ["Номер мерчанта", "Дата зачисления"], right_on=["Номер мерчанта", "Дата"],
          suffixes=('_sber', '_bank'))
# проверка на пустые значения после слияния
for i in range(1, len(itog)):
    if pd.isna(itog.loc[i, "Получено"]):
        itog.loc[i, "Получено"] = 0
        itog.loc[i, "Возврат1"] = 0
        itog.loc[i, "Возврат2"] = 0
itog["Проверка"] = itog["Сумма операции"] - itog["Получено"] - itog["Возврат1"] - itog["Возврат2"]
itog["Накопительно"] = 0
itog.loc[1, "Накопительно"] =  itog.loc[1, "Проверка"]
for i in range(2, len(itog)):
    if itog.loc[i, "Номер мерчанта"] == itog.loc[i-1, "Номер мерчанта"]:
        itog.loc[i, "Накопительно"] = itog.loc[i-1, "Накопительно"] + itog.loc[i, "Проверка"]
    else:
        itog.loc[i, "Накопительно"] = itog.loc[i, "Проверка"]

# print (itog)
# exit("my stop")

# для записи страницы с форматированием дат
def write_worksheet(panda_df, name_worksheet):
    panda_df.to_excel(file_name, sheet_name=name_worksheet, index=False)
    worksheet = writer.sheets[name_worksheet]
    count = 0
    for col in panda_df.columns:
        if "ДАТА" in col.upper():
            worksheet.set_column(count, count, 15, format_data)
        count += 1


# запись данных в эксель
writer = pd.ExcelWriter(filename_xls, engine='xlsxwriter')
workbook  = writer.book
format_data = workbook.add_format({'num_format': 'dd/mm/yy'})
with writer as file_name:
    write_worksheet(df, "Вся выписка")
    write_worksheet(svod, "Эквайринг банк")
    write_worksheet(svod_sber_df, "Эквайринг сбер")
    write_worksheet(itog, "Проверка")



