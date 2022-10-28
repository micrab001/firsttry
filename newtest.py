# import pandas as pd

time_delt = 3
time_wrong = "19:22:12"
time_new = int(time_wrong[:2])+time_delt
if time_new >= 24:
    raise ("ошибка конвертации времени")
time_new = str(time_new) + time_wrong[-6:]

print(time_new)

# full_data = pd.DataFrame({'Сумма транзакции': [1894, 511, 11, 1493, 36],
# 'Тип операции': ["Покупка", "Возврат", "Покупка", "Возврат", "Покупка"]})
#
# print(full_data)
#
# for i in range(1, len(full_data)):
#     if full_data.loc[i, 'Тип операции' ] == "Возврат":
#         full_data.loc[i, 'Сумма транзакции' ] = full_data.loc[i, 'Сумма транзакции' ] * -1
#
# print(full_data)