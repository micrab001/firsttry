import requests
import socket
comp_ip = socket.gethostbyname(socket.gethostname())
localnet = comp_ip[0:comp_ip.rfind(".")-len(comp_ip)+1]

rez = []
# rez1 = []
for i in range(1,256):
    net_address = f"http://{localnet}{str(i)}"
    try:
        responce = requests.get(net_address, timeout=(0.01, 1))
        print(f"Получение ответа от адреса {net_address} : ", responce)
        rez.append(f"Получен ответа от адреса {net_address}: {responce}")
        # rez1.append(socket.gethostbyaddr(f"{localnet}{str(i)}"))
    except requests.exceptions.ConnectionError:
        print(f"Получение ответа от адреса {net_address} : Ошибка! Подключение не установлено")

print(rez)
# print(rez1)



# import json
#
# link = "http://192.168.50.156/YamahaExtendedControl"
#
# # link += "/v2/system/getDeviceInfo"
# getDeviceInfo = {
#     "response_code": 0,
#     "model_name": "RX-V485",
#     "destination": "F",
#     "device_id": "F086204A289E",
#     "system_id": "084BED33",
#     "system_version": 1.78,
#     "api_version": 2.11,
#     "netmodule_generation": 2,
#     "netmodule_version": "1107    ",
#     "netmodule_checksum": "79DE0042",
#     "serial_number": "Y219680ZP",
#     "category_code": 1,
#     "operation_mode": "normal",
#     "update_error_code": "00000000",
#     "net_module_num": 1,
#     "update_data_type": 0
# }
#
# # "power": "standby"
#
# link += "/v2/main/setPower?power=standby"
# responce = requests.get(link, timeout=10)
# if responce.status_code == 200:
#     # for el in responce.json()
#     print(json.dumps(responce.json(), indent= 4))
# else:
#     print("Получение списка ККТ, ответ: ", responce)


# link = "http://192.168.50.156"
# responce = requests.get(link, timeout=10)
# if responce.status_code == 200:
#     # for el in responce.json()
#     print(json.dumps(responce.json(), indent= 4))
# else:
#     print("Получение списка ККТ, ответ: ", responce)
