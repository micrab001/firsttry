import pandas as pd
from tkinter import filedialog
import os
import requests

# выбор нужной директории и создание списка файлов в нем с расширением тхт
dirname = filedialog.askdirectory(initialdir=os.getcwd()).replace("/", chr(92))
all_dir = os.listdir(dirname)
filesnames = [dirname+chr(92)+f for f in all_dir if os.path.isfile(dirname+chr(92)+f) and ".txt" in f]

# задание некоторых начальных переменных
filename_xls = dirname+chr(92)+"all_data.xlsx"

data_table = {"Дата":[], "время":[], "loc_ip":[], "loc_port":[], "remote_ip":[], "remote_port":[], "byte_in":[], "byte_out":[]}

for filename in filesnames:
    with open(filename, "r") as data_file:
        line = data_file.readline()
        for line in data_file:
            line_list = line.split("\t")
            line_list.insert(1, line_list[0].split(" ")[1])
            line_list[0] = line_list[0].split(" ")[0]
            line_list[-1] = line_list[-1].replace("\n", "")
            for i in range(len(line_list)):
                if line_list[i].isdigit():
                    line_list[i] = int(line_list[i])
            data_table["Дата"].append(line_list[0])
            data_table["время"].append(line_list[1])
            data_table["loc_ip"].append(line_list[2])
            data_table["loc_port"].append(line_list[3])
            data_table["remote_ip"].append(line_list[4])
            data_table["remote_port"].append(line_list[5])
            data_table["byte_in"].append(line_list[6])
            data_table["byte_out"].append(line_list[7])
    print(f"Обработка файла {filename}")

data ={'status': [],
       'country': [],
       'countryCode': [],
       'region': [],
       'regionName': [],
       'city': [],
       'zip': [],
       'lat': [],
       'lon': [],
       'timezone': [],
       'provider': [],
       'organization': [],
       'as': [],
       'query IP': []}

def get_info_by_ip(ip="127.0.0.1"):
    try:
        response = requests.get(url = f"http://ip-api.com/json/{ip}", timeout=10).json()
        data['status'].append(response.get('status'))
        data['country'].append(response.get('country'))
        data['countryCode'].append(response.get('countryCode'))
        data['region'].append(response.get('region'))
        data['regionName'].append(response.get('regionName'))
        data['city'].append(response.get('city'))
        data['zip'].append(response.get('zip'))
        data['lat'].append(response.get('lat'))
        data['lon'].append(response.get('lon'))
        data['timezone'].append(response.get('timezone'))
        data['provider'].append(response.get('isp'))
        data['organization'].append(response.get('org'))
        data['as'].append(response.get('as'))
        data['query IP'].append(response.get('query'))
    except:
        print("!! Please check your connection !!")


df = pd.DataFrame(data_table)
print("Датафрейм создан")
df["byte_all"] = df["byte_in"]+ df["byte_out"]
print("Поле суммы байт добавлено")
svod_df = df[["remote_ip", "byte_in", "byte_out", "byte_all"]].groupby("remote_ip", as_index=False).sum()
print("Сгрупированно по IP ")
svod_df = svod_df[svod_df["byte_all"] > 100000000]
print("Выборка по большим IP ")
svod_df = svod_df.sort_values(by="byte_all")
print("Сортировка по IP ")

for i in range(len(svod_df)):
    print(svod_df.iloc[i, 0])
    get_info_by_ip(svod_df.iloc[i, 0])

df_ip_ifo = pd.DataFrame(data)

def write_worksheet(panda_df, name_worksheet):
    panda_df.to_excel(file_name, sheet_name=name_worksheet, index=False)
    worksheet = writer.sheets[name_worksheet]
    count = 0
    for col in panda_df.columns:
        if "ДАТА" in col.upper():
            worksheet.set_column(count, count, 15, format_data)
        count += 1
    print("формирование листа ", name_worksheet)

writer = pd.ExcelWriter(filename_xls, engine='xlsxwriter')
workbook  = writer.book
format_data = workbook.add_format({'num_format': 'dd/mm/yy'})
print("старт записи в эксель")
with writer as file_name:
    write_worksheet(df, "Все данные")
    write_worksheet(svod_df, "big traffic ip")
    write_worksheet(df_ip_ifo, "ip info")
print("запись завершена")

