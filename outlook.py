# from datetime import datetime, date, timedelta
import win32com.client
from tkinter import filedialog

class Work_with_outlook():

    def __init__(self):
        outlook = win32com.client.Dispatch('outlook.application')
        mapi = outlook.GetNamespace("MAPI")
        # так можно посмотреть все эккаунты, настроенные в аутлук
        # for account in self.mapi.Accounts:
        #     print(account.DeliveryStore.DisplayName)
        # так можно посмотреть все папки в эккаунте
        # for idx, folder in enumerate(self.mapi.Folders("partneryr001@outlook.com").Folders):
        #     print(idx + 1, folder)
        # создаем объект всех сообщений из эккаунта из определенной папки (в нашем случае "входящие")
        self.messages = mapi.Folders("partneryr001@outlook.com").Folders(2).Items
        # так можно перебрать все сообщения по одному (в данном случае берем 10 сообщений)
        # for message in list(messages)[:10]:
        #     print(message.Subject, message.ReceivedTime, message.SenderEmailAddress)
        # а так можно менять дату
        # today = datetime.today()
        # # first day of the month
        # start_time = today.replace(day=1, hour=0, minute=0, second = 0).strftime('%Y-%m-%d %H:%M %p')
        # # today 12am
        # end_time = today.replace(hour=0, minute=0, second=0).strftime('%Y-%m-%d %H:%M %p')

        datefrom = "01/09/22" # брать с 01 числа
        dateto = "02/11/22" # берется диапазон включая крайние даты плюс 1 день!

        ini_dir = "d:\\OneDrive\\Рабочие документы\\Эквайринг\\"
        self.process_sbp = True
        self.process_card = True
        chk_dir = "не выбран каталог, отмена программы"
        self.dir_sbp = filedialog.askdirectory(initialdir=ini_dir, title="Выбор каталога для файлов СБП").replace("/", chr(92))
        ini_dir = "\\".join(self.dir_sbp.split("\\")[:-1])
        if self.dir_sbp == "":
            self.process_sbp = False
        self.dir_card = filedialog.askdirectory(initialdir=ini_dir, title="Выбор каталога для файлов эквайринга").replace("/", chr(92))
        if self.dir_card == "":
            self.process_card = False
        self.messages = self.set_messages_filter("[ReceivedTime] >= '" + datefrom + "' And [ReceivedTime] <= '" + dateto + "'")

    def set_messages_filter(self, flt: str = ""):
        return self.messages.Restrict(flt)


def email_filter(flt_email):
    return "[SenderEmailAddress] = '" + flt_email + "'"


def messages_count(msg, stroka = ""):
    print(f"Сообщений {stroka}: {len(msg)}")


def messages_list_print(msg):
    for message in list(msg):
        print(message.SenderEmailAddress, message.ReceivedTime.strftime('%d_%m_%Y_%H_%M_%S'), message.Subject)

def message_attachments(msg, outputdir):
    count = 1
    try:
        for message in list(msg):
            print(message.SenderEmailAddress, message.ReceivedTime.strftime('%d_%m_%Y_%H_%M_%S'), message.Subject)
            try:
                for attachment in message.Attachments:
                    attachment.SaveAsFile(f"{outputdir}\\{count:04} {message.ReceivedTime.strftime('%d_%m_%Y_%H_%M_%S')} {attachment.FileName}")
                    print(f"attachment {attachment.FileName} from {message.sender} saved")
                    count += 1
            except Exception as e:
                print("error when saving the attachment:" + str(e))
    except Exception as e:
        print("error when processing emails messages:" + str(e))
    print(f"записано {count-1} вложений")

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    x = Work_with_outlook()
    messages_count(x.messages, "всего")
    if x.process_sbp:
        mes_sbp = x.set_messages_filter(email_filter("reestr-sbp@alfabank.ru"))
        messages_count(mes_sbp, "от СБП")
        messages_list_print(mes_sbp)
        message_attachments(mes_sbp, x.dir_sbp)

    mes_card = x.set_messages_filter(email_filter("Esupport@alfa-bank.info"))
    if x.process_card:
        messages_count(mes_card, "эквайринг")
        messages_list_print(mes_card)
        message_attachments(mes_card, x.dir_card)




