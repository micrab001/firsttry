import pandas as pd
import json
from tkinter import filedialog
import os
import datetime

# выбор нужной директории и создание списка файлов в нем с расширением тхт

dirname = filedialog.askdirectory(initialdir="d:\\OneDrive\\Рабочие документы\\Выписки Альфа\\").replace("/", chr(92)) #initialdir=os.getcwd()
all_dir = os.listdir(dirname)
filesnames = [dirname+chr(92)+f for f in all_dir if os.path.isfile(dirname+chr(92)+f) and ".txt" in f]

# задание некоторых начальных переменных
all_data = ""
filename_xls = dirname+chr(92)+"all_data_"+dirname.split("\\")[-1]+".xlsx"
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
                                    f'"Возврат2":{tmp_lst[tmp_lst.index("ПОКУПКИ")+1].split("/")[1].rstrip(".").replace(",", "").replace(".НДС","")}, '

all_data = all_data.replace(", }", "}")
all_data = all_data[0: all_data.rfind(",")]
all_data = "[" + all_data + "]"

with open("test.json", "w") as tst_file:
    tst_file.write(all_data)
print("Выписки обработаны")

# начало работы с данными
# считывание выписок
df = pd.DataFrame(json.loads(all_data))
# создание выборки данных из выписок по эквайрингу
svod = df[["Дата", "Сумма", "ДатаПоступило", "Получатель", "НазначениеПлатежа", "Номер мерчанта",
           "Сумма комиссии", "Возврат1", "Возврат2"]] # "Дата операции магазин"
svod = svod[svod["Номер мерчанта"].notnull()]
svod = svod.sort_values(by=["Номер мерчанта", "Дата"], ignore_index = True)
svod["Получено"] = svod["Сумма"] + svod["Сумма комиссии"]
svod["Получено"] = svod["Получено"].astype("int")
svod["Учтено"] = "Нет"
svod.reset_index(drop=True, inplace=True)

print("таблица выписок сформирована")

# считываение данных по эквайрингу от Сбера
filename = filedialog.askopenfilename(initialdir="d:\\OneDrive\\Рабочие документы\\Эквайринг\\") #initialdir=os.getcwd()
sber_df = pd.read_excel(pd.ExcelFile(filename), "Sheet0")
print("чтение данных Сбер завершено")
sber_df["Дата зачисления"] = sber_df["Дата выгрузки в АБС"].apply(convert_data)
sber_df["Дата зачисления"] = sber_df["Дата зачисления"].astype("str").apply(str_to_data)
sber_df["Дата операции магазин"] = sber_df["Дата операции магазин"].apply(convert_data)
sber_df["Дата операции магазин"] = sber_df["Дата операции магазин"].astype("str").apply(str_to_data)
# Номер мерчанта, Дата операции магазин, Дата выгрузки в АБС, Сумма операции, Сумма комиссии, Сумма расчета
sber_df["Номер мерчанта"] = sber_df["Номер мерчанта"].astype("str")
# получение выборки данных из эквайринга Сбера
svod_sber_df = sber_df[["Номер мерчанта", "Дата зачисления", "Сумма операции", "Сумма комиссии", "Сумма расчета"]]
svod_sber_df = svod_sber_df.groupby(["Номер мерчанта", "Дата зачисления"], as_index=False).sum() #[["Сумма операции"]]
svod_sber_df.reset_index(drop=True, inplace=True)
# дополнительный расчет по сберу
sber_test = sber_df[["Номер мерчанта", "Дата зачисления", "Дата операции магазин", "Сумма операции", "Сумма комиссии", "Сумма расчета"]]
sber_test = sber_test.groupby(["Номер мерчанта", "Дата зачисления", "Дата операции магазин"], as_index=False).sum()

# слияние данных сбера и банка
print("Слияние данных и поиск недостающих платежей")
itog = svod_sber_df.merge(svod, how = "left", left_on= ["Номер мерчанта", "Сумма операции"], right_on=["Номер мерчанта", "Получено"],
          suffixes=('_sber', '_bank'), validate="one_to_one")

# проверка на потерянные платежи при банковском переводе и не учтенные при проверке сумм зачислений по эквайрингу в середине месяца
new_line = {'Номер мерчанта':[], 'Дата зачисления':[], 'Сумма операции':[], 'Сумма комиссии_sber':[], 'Сумма расчета':[], 'Дата':[], 'Сумма':[],
          'ДатаПоступило':[], 'Получатель':[], 'НазначениеПлатежа':[], 'Сумма комиссии_bank':[], 'Возврат1':[], 'Возврат2':[], 'Получено':[]}

# первый проход по переводам эквайринга и их первичная проверка
# если совпадение с платежом, добавляем в таблицу банковских платежей, что он найден

svod_sber_df[["Дата", "Сумма", "ДатаПоступило", "Получатель", "НазначениеПлатежа",
           "Сумма комиссии банк", "Возврат1", "Возврат2", "Получено"]] = ""
for i in range(0, len(svod_sber_df)):
    find_payment = svod.loc[(svod["Номер мерчанта"] == svod_sber_df.loc[i, "Номер мерчанта"]) & (svod["Дата"] >= svod_sber_df.loc[i, "Дата зачисления"])
                            & (svod["Получено"] == svod_sber_df.loc[i, "Сумма операции"])]
    if len(find_payment) == 1:
        find_payment_idx = find_payment.index[0]
        svod_sber_df.loc[i, "Дата"] = find_payment.loc[find_payment_idx, "Дата"]
        svod_sber_df.loc[i, "Сумма"] = find_payment.loc[find_payment_idx, "Сумма"]
        svod_sber_df.loc[i, "ДатаПоступило"] = find_payment.loc[find_payment_idx, "ДатаПоступило"]
        svod_sber_df.loc[i, "Получатель"] = find_payment.loc[find_payment_idx, "Получатель"]
        svod_sber_df.loc[i, "НазначениеПлатежа"] = find_payment.loc[find_payment_idx, "НазначениеПлатежа"]
        svod_sber_df.loc[i, "Сумма комиссии банк"] = find_payment.loc[find_payment_idx, "Сумма комиссии"]
        svod_sber_df.loc[i, "Возврат1"] = find_payment.loc[find_payment_idx, "Возврат1"]
        svod_sber_df.loc[i, "Возврат2"] = find_payment.loc[find_payment_idx, "Возврат2"]
        svod_sber_df.loc[i, "Получено"] = find_payment.loc[find_payment_idx, "Получено"]
        svod.loc[find_payment_idx, "Учтено"] = "Да"
    elif len(find_payment) == 0:
        svod_sber_df.loc[i, "Сумма"] = 0
    else:
        svod_sber_df.loc[i, "Сумма"] = -1

# далее идет подбор не найденных операций и платежей


def find_equal_transaction (sber_oper :list, bank_oper :list) -> dict:
    """ Подпрограмма поиска совпадений по суммам, на входе два листа
    sber_oper - список с цифрами, содержащий хотя бы 1 элемент, это не найденные операции Сбербанка эквайринг
    bank_oper - список с цифрами, содержащий хотя бы 1 элемент, это не найденные операции Банка по переводам
    oper_calc - возвращает словарь подобранных совпадений, где номер элемента листа по ключу банка соответствует
                совпадению с тем же номером элемента листа по ключу сбера
    """

    oper_calc = {"sber": [], "bank": []}
    naideno = []
    for i in range(len(sber_oper)):
        if sber_oper[i] in naideno:
            continue
        for j in range(len(bank_oper)):
            if sber_oper[i] in naideno or bank_oper[j] in naideno:
                continue
            sum_op = sber_oper[i]
            sum_bk = bank_oper[j]
            found_flag = True
            op = i
            bk = j
            oper_calc_tmp_sber = [sber_oper[i]]
            oper_calc_tmp_bank = [bank_oper[j]]
            while found_flag:
                if sber_oper[i] in naideno:
                    op += 1
                    continue
                if bank_oper[j] in naideno:
                    bk += 1
                    continue
                if sum_op < sum_bk:
                    op += 1
                    if op > len(sber_oper) - 1:
                        break
                    else:
                        sum_op += sber_oper[op]
                        oper_calc_tmp_sber.append(sber_oper[op])
                elif sum_op > sum_bk:
                    bk += 1
                    if bk > len(bank_oper) - 1:
                        break
                    else:
                        sum_bk += bank_oper[bk]
                        oper_calc_tmp_bank.append(bank_oper[bk])
                else:
                    found_flag = False
                    oper_calc["sber"].append(oper_calc_tmp_sber)
                    oper_calc["bank"].append(oper_calc_tmp_bank)
                    naideno += oper_calc_tmp_sber + oper_calc_tmp_bank


    print(oper_calc)
    print(naideno)
    return oper_calc

# берем мерчанты с ненайденными платежами
merch_list = set(svod_sber_df[svod_sber_df["Сумма"] == 0]["Номер мерчанта"].tolist())
if len(merch_list) > 0:
    # запускаем цикл по мерчантам
    for one_merch in merch_list:
        # берем срез данных по эквайрингу по мерчанту и не найденной сумме и по необработанным платежам банка
        sber_ne_naiden = svod_sber_df[(svod_sber_df["Номер мерчанта"] == one_merch) & (svod_sber_df["Сумма"] == 0)]
        find_payment = svod.loc[(svod["Номер мерчанта"] == one_merch) & (svod["Учтено"] == "Нет")] # (svod["Дата"] >= sber_ne_naiden.iloc[0]["Дата зачисления"])
        if len(find_payment) > 0:
            print(one_merch)
            match_found = find_equal_transaction(sber_ne_naiden["Сумма операции"].tolist(), find_payment["Получено"].tolist())
            for i in range(len(match_found["sber"])):
                for j in range(len(match_found["sber"][i])):
                    svod_sber_df.loc[(svod_sber_df["Номер мерчанта"] == one_merch) & (svod_sber_df["Сумма"] == 0) &
                                     (svod_sber_df["Сумма операции"] == match_found["sber"][i][j] ), ["НазначениеПлатежа"]] = f'{match_found["bank"][i]}={sum(match_found["bank"][i])}'
            for i in range(len(match_found["bank"])):
                for j in range(len(match_found["bank"][i])):
                    svod.loc[(svod["Номер мерчанта"] == one_merch) & (svod["Учтено"] == "Нет") &
                                     (svod["Получено"] == match_found["bank"][i][j] ), ["Учтено"]] = "Да"





# {'sber': [[5549, 43031], [75572, 15907], [50326]], 'bank': [[48580], [63469, 28010], [33864, 16462]]}

# проход по банковским платежам и ищем, попали ли они в проверку
# если не попали, добавляем их в таблицу
# for i in range(1, len(svod)):
#     find_lost_payment = itog.loc[itog["Номер мерчанта"] == svod.loc[i, "Номер мерчанта"]]
#     if len(find_lost_payment) == 0:
#         print("**************** потерянный мерчант (ищи в выписках)", svod.loc[i, "Номер мерчанта"])
#         continue
#     tmp_list_data = sorted(find_lost_payment["Дата"].tolist()) # ограничения по датам эквайринга
#     end_data = tmp_list_data[-1] # последняя дата эквайринга
#     start_data = tmp_list_data[0] # первая дата (как бы ограничиваем месяц)
#     find_lost_payment = find_lost_payment.loc[find_lost_payment["Дата зачисления"] == svod.loc[i, "Дата"]]
#     find_lost_payment = find_lost_payment.loc[find_lost_payment["Получено"] == svod.loc[i, "Получено"]]
#     if len(find_lost_payment) == 0 and svod.loc[i, "Дата"] <= end_data and svod.loc[i, "Дата"] >= start_data:
#         new_line['Номер мерчанта'].append(svod.loc[i, "Номер мерчанта"])
#         new_line['Дата зачисления'].append(svod.loc[i, "Дата"])
#         new_line['Сумма операции'].append(0)
#         new_line['Сумма комиссии_sber'].append(0)
#         new_line['Сумма расчета'].append(0)
#         new_line["Дата"].append(svod.loc[i, "Дата"])
#         new_line["Сумма"].append(svod.loc[i, "Сумма"])
#         new_line["ДатаПоступило"].append(svod.loc[i, "ДатаПоступило"])
#         new_line["Получатель"].append(svod.loc[i, "Получатель"])
#         new_line["НазначениеПлатежа"].append(svod.loc[i, "НазначениеПлатежа"])
#         new_line["Сумма комиссии_bank"].append(svod.loc[i, "Сумма комиссии"])
#         new_line["Возврат1"].append(svod.loc[i, "Возврат1"])
#         new_line["Возврат2"].append(svod.loc[i, "Возврат2"])
#         new_line["Получено"].append(svod.loc[i, "Получено"])
# if len(new_line['Номер мерчанта']) > 0:
#     new_line_df = pd.DataFrame(new_line)
#     itog = pd.concat([itog, new_line_df], ignore_index = True)
#     itog = itog.sort_values(by = ["Номер мерчанта",	"Дата зачисления"], ignore_index = True)
# print("Недостающие платежи проверены")
# проверка на пустые значения после слияния
for i in range(1, len(itog)):
    if pd.isna(itog.loc[i, "Получено"]):
        itog.loc[i, "Получено"] = 0
        itog.loc[i, "Возврат1"] = 0
        itog.loc[i, "Возврат2"] = 0
        itog.loc[i, "Сумма комиссии_bank"] = 0
        itog.loc[i, "Сумма"] = 0

itog["Проверка"] = itog["Сумма операции"] - itog["Получено"] # - itog["Возврат1"] - itog["Возврат2"]
itog["Накопительно"] = 0
itog.loc[1, "Накопительно"] =  itog.loc[1, "Проверка"]
for i in range(2, len(itog)):
    if itog.loc[i, "Номер мерчанта"] == itog.loc[i-1, "Номер мерчанта"]:
        itog.loc[i, "Накопительно"] = itog.loc[i-1, "Накопительно"] + itog.loc[i, "Проверка"]
    else:
        itog.loc[i, "Накопительно"] = itog.loc[i, "Проверка"]

# собирает итоги месяца по мерчанту, если проверка не равна 0 надо искать потерянные платежи
itog_sum = itog[["Номер мерчанта", "Проверка"]]
itog_sum = itog_sum.groupby(["Номер мерчанта"], as_index=False).sum()
itog_sum["Финальная проверка"] = itog_sum["Проверка"].abs()
# поиск не найденных платежей по совпадению сумм
# поиск по расхождениям в месяце и корректировка предыдущего списка итогов месяца
list_plat = list(map(abs,itog_sum[itog_sum["Проверка"]!= 0]["Проверка"].tolist()))
find_plat =  svod[svod["Получено"].isin(list_plat)]
list_plat = find_plat["Получено"].tolist()
for i in range(1, len(itog_sum)):
    if itog_sum.loc[i, "Финальная проверка"] in list_plat:
        itog_sum.loc[i, "Финальная проверка"] = 0
#
#
print("расчеты произведены")
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
    print("формирование листа ", name_worksheet)

# запись данных в эксель
writer = pd.ExcelWriter(filename_xls, engine='xlsxwriter')
workbook  = writer.book
format_data = workbook.add_format({'num_format': 'dd/mm/yy'})
with writer as file_name:
    write_worksheet(df, "Вся выписка")
    write_worksheet(svod, "Эквайринг банк")
    write_worksheet(svod_sber_df, "Эквайринг сбер")
    write_worksheet(itog, "Проверка")
    write_worksheet(itog_sum, "Проверка месяц")
    write_worksheet(find_plat, "Не найденные платежи")
    write_worksheet(sber_test, "Сбер операции в магазине")
print("запись завершена")




