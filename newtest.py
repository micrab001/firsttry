import sqlite3

import pandas as pd

full_data = pd.read_excel("d:\\OneDrive\\Рабочие документы\\Эквайринг\\2022_11\\full_data.xlsx", "Sheet0")

full_data.insert(full_data.columns.get_loc('Банк') + 1, 'Наименование банка', "")
full_data["Банк"] = full_data["Банк"].astype("str")

db = sqlite3.connect("c:\\Vrem\\Python10\\bd_bin_code1.db")
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

for i in range(0, len(full_data)):
    if len(full_data.loc[i, "Банк"]) > 7 and full_data.loc[i, "Система"] == "СБП":
        full_data.loc[i, "Наименование банка"] = convert_bic_code(full_data.loc[i, "Банк"][:-2])
    else:
        full_data.loc[i, "Наименование банка"] = convert_bin(full_data.loc[i, "Карта или счет"])
        if full_data.loc[i, "Банк"] == "0" and "АЛЬФА" not in full_data.loc[i, "Наименование банка"].upper():

            pass





full_data.to_excel("d:\\OneDrive\\Рабочие документы\\Эквайринг\\2022_11\\full_data1.xlsx", sheet_name="Sheet0", index=False)
db.close()