"""
Сделать связку документа .xls (поднять на локальной машине) и таблицы Google Spreadsheet (использовать личный аккаунт).
Обновление связей по триггеру (временной) или по запросу. Данные из .xls копируются в гугл таблицу на заданный лист.
Данные в таблице - массив ячеек 10х1000 со случайными символами.

Ссылка на таблицу - https://docs.google.com/spreadsheets/d/1jHcxOKB3uyXSaql7_IthROIAs-Vsf9iA4-sRzHCWY48/
"""

from pprint import pprint
import httplib2
import apiclient
from oauth2client.service_account import ServiceAccountCredentials
import xlrd
from threading import Thread
from time import sleep

CREDENTIALS_FILES = "creds.json"
spreadsheets_id = "1jHcxOKB3uyXSaql7_IthROIAs-Vsf9iA4-sRzHCWY48"
sheet = "Лист1"
update_time = 5  # в секундах
stop = False


# считывание данных из файла excel
def excel_data():
    rb = xlrd.open_workbook('file.xls')
    sheet = rb.sheet_by_index(0)
    data = []
    for row in range(sheet.nrows):
        data.append(sheet.row_values(row))
    return data


# обновление google sheets
def update():
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        CREDENTIALS_FILES,
        ["https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive"]
    )

    service = apiclient.discovery.build("sheets", "v4", http=credentials.authorize(httplib2.Http()))

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheets_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": "{}!A1:J1000".format(sheet),
                 "majorDimension": "ROWS",
                 "values": excel_data()}]}).execute()


def periodical_update():
    global stop
    while not stop:
        update()
        # print("-- фоновое обновление --")
        sleep(update_time)


def main():
    global stop
    while not stop:
        a = input("Если вы хотите чтобы данные обновились прям сейчас введите now. \n"
                  "Если хотите обновить время обновления, то введите число секунд (целое). \n"
                  "Если хотите указать другой лист напишите 'sheet:ИМЯЛИСТА'\n"
                  "Если хотите выйти из программы введите exit\n"
                  "Ввод: ")
        if a.lower() == "now":
            update()
            print("Обновлено\n")
        elif a.isnumeric():
            global update_time
            update_time = int(a)
            print("Время обновленияч уствновлено на {} минут\n".format(update_time))
        elif a.lower() == "exit":
            stop = True
        else:
            global sheet
            sheet = a[6:]
            print("Установлен лист: {}\n".format(sheet))


if __name__ == "__main__":
    periodical_update = Thread(target=periodical_update)
    main = Thread(target=main)
    periodical_update.start()
    main.start()
    periodical_update.join()
    main.join()

