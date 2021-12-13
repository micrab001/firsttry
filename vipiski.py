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
strt_time = datetime.date(1900, 1, 1) # объект для конвертации даты

# конвертация даты из строки в число в формат для Excel
def str_to_data(ds): # предполагается, что формат даты д-м-г если ошибка то берет г-м-d
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
            return (datetime.date(dd[2] if dd[2]>100 else dd[2]+2000, dd[1], dd[0]) - strt_time).days + 2
        except ValueError:
            return (datetime.date(dd[0] if dd[0]>100 else dd[0]+2000, dd[1], dd[2]) - strt_time).days + 2


for filename in filesnames:
    with open(filename, "r") as data_file:
        flag = False
        for line in data_file:
            line = line.replace("\n", "").replace('"', "'").replace("\\", "/")
            if line in ("СекцияДокумент=Платежное поручение", "СекцияДокумент=Банковский ордер", "СекцияДокумент=Платежное требование"):
                all_data += "{"
                flag = True
            if "КонецДокумента" in line:
                all_data += "}, \n"
                flag = False
            else:
                if flag:
                    tmp_list = line.split("=")
                    tmp_list[1] = tmp_list[1].upper()
                    if "Дата" in tmp_list[0]:
                        if tmp_list[1] != "":
                            tmp_list[1] = str_to_data(tmp_list[1])
                            # tmp_list[1] = convert_data(tmp_list[1])
                        # tmp_list[1] = tmp_list[1].replace(".", "/")
                    if "Сумма" == tmp_list[0] or "Дата" in tmp_list[0] and tmp_list[1] != "":
                        all_data += f'"{tmp_list[0]}":{tmp_list[1]}, '
                    else:
                        all_data += f'"{tmp_list[0]}":"{tmp_list[1]}", '
                    if tmp_list[0] == "НазначениеПлатежа" and "ЗАЧИСЛЕНИЕ СРЕДСТВ ПО ОПЕРАЦИЯМ С МБК (НА ОСНОВАНИИ РЕЕСТРОВ ПЛАТЕЖЕЙ)." in tmp_list[1]:
                        tmp_lst = tmp_list[1].split()
                        all_data += f'"Номер мерчанта":"{tmp_lst[11].strip("№.")}", "Дата операции магазин":"{str_to_data(tmp_lst[14].rstrip("."))}", ' \
                                    f'"Сумма комиссии":{tmp_lst[16].rstrip(".").replace(",", "")}, "Возврат1":{tmp_lst[19].split("/")[0].replace(",", "")}, ' \
                                    f'"Возврат2":{tmp_lst[19].split("/")[1].rstrip(".").replace(",", "")}, '

all_data = all_data.replace(", }", "}")
all_data = all_data[0: all_data.rfind(",")]
all_data = "[" + all_data + "]"

with open("test.json", "w") as tst_file:
    tst_file.write(all_data)

df = pd.DataFrame(json.loads(all_data))
df.to_excel(filename_xls, index=False)


