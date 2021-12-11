import pandas as pd
import json
from tkinter import filedialog
from os import getcwd

filename = filedialog.askopenfilename(initialdir=getcwd())

filename_json = filename.removesuffix(".txt")+".new"
filename_xls = filename.removesuffix(".txt")+".xlsx"


with open(filename_json, "w") as data_out_file:
    # print("[", file=data_out_file, end="")
    with open(filename, "r") as data_file:
        flag = False
        for line in data_file:
            line = line.replace("\n", "")
            line = line.replace('"', "'")
            if "СекцияДокумент=Платежное поручение" in line or "СекцияДокумент=Банковский ордер" in line :
                print("{", file=data_out_file, end="")
                flag = True
            if "КонецДокумента" in line:
                print("}, ", file=data_out_file)
                flag = False
            else:
                if flag:
                    tmp_list = line.split("=")
                    print(f'"{tmp_list[0]}":"{tmp_list[1]}",', file=data_out_file, end=" ")


with open(filename_json, "r") as read_file:
    all_data = read_file.read()

all_data = all_data.replace(", }", "}")
all_data = all_data[0: all_data.rfind(",")]
all_data = "[" + all_data + "]"

df = pd.DataFrame(json.loads(all_data))
df.to_excel(filename_xls)


