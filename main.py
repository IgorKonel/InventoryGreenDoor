"""Библиотека для получения расположения файла"""
import os

import creds  # Импорт библиотеки, содержащей API key

from googleapiclient.discovery import build  # Библиотеки для работы google sheets API

import openpyxl  # Библиотека для работы с Excel
from openpyxl.styles import Font, Alignment

NAME_LIST = "Invent WD/ZD"


def print_list(list_resp, str_resp):
    print(f'\n{str_resp} \n')
    for pos, num in list_resp:
        print(f'{pos:<60} {num:>8}')


def print_dict_list(list_resp, str_resp):
    print(f'\n{str_resp} \n')
    for pos in list_resp:
        print(f'{pos[0]:<60} {pos[1]:<9} {pos[2]:<11} {pos[3]:>3}')


# def get_service_sacc():
#     """
#     Чтение таблиц, к которым выдан доступ
#
#     godinventory@inventory-328117.iam.gserviceaccount.com
#
#     :return:
#     """
#     creds_json = os.path.dirname(__file__) + "/creds/sacc1.json"
#     print(creds_json)
#     scopes = ['https://www.googleapis.com/auth/spreadsheets']
#
#     creds_service = ServiceAccountCredentials.from_json_keyfile_name(creds_json, scopes).authorize(httplib2.Http())
#     return build('sheets', 'v4', http=creds_service)


def get_service_simple():
    """
    Чтение таблиц без возможности их редактирования
    :return:
    """
    return build('sheets', 'v4', developerKey=creds.api_key)


def get_resp_several(range_list):
    path = NAME_LIST + '!' + range_list
    # print('path = \'', path, '\'',  sep='')
    return sheet.values().get(spreadsheetId=sheet_id, range=path).execute()  # ['valueRanges'][0]['values']


def get_resp(range_list):
    path = NAME_LIST + '!' + range_list
    # print('path =', path)
    return sheet.values().batchGet(spreadsheetId=sheet_id, ranges=[path]).execute()['valueRanges'][0]['values']


def get_list_order(range_list):
    return [s for s in range_list if s[len(range_list[0]) - 1] != '0']


service = get_service_simple()  # С использованием API
# service = get_service_sacc()  # С использованием OAuth 2.0 Client IDs
sheet = service.spreadsheets()

# TODO: Сделать возможность вставлять ID в оконном приложении
sheet_id = '13nzuCdZrcOaQMBqTKMJOzehQOff97ixVoR5opfBfBjM'

"""
************************************************************************************************************************
                                             Запрос на получение данных
************************************************************************************************************************
"""

""" 
------------------------------------------------------------------------------------------------------------------------
                                                 Для одного диапазона
------------------------------------------------------------------------------------------------------------------------
"""
# TODO:
resp_house = get_resp('AP5:AQ48')  # Хоз-товары на всех
resp_tea = get_resp('AP49:AQ85')  # Чаи на всех
resp_art_count = get_resp('B5:C48')  # Артикулы и кол-во

for i in range(len(resp_house)):
    resp_house[i].insert(1, resp_art_count[i][0])
    resp_house[i].insert(2, resp_art_count[i][1])

# print_list(resp_house, 'Список общих хоз-товаров:')
# print_list(resp_tea, 'Список общих чаёв:')
# print_list(resp_art_count, 'Список артикулов и кол-ва')

# TODO:
order_house = get_list_order(resp_house)
order_tea = get_list_order(resp_tea)
# print_dict_list(order_house, 'Хозы на заказ:')
# print_list(order_tea, 'Чай на заказ:')

# resp_double = sheet.values().batchGet(spreadsheetId=sheet_id, ranges=["Invent WD/ZD!AP49:AP70",
# "Invent WD/ZD!AR49:AR70"]).execute()['valueRanges'][0]  #[0]['values']

# resp = sheet.values().get(spreadsheetId=sheet_id, range="Invent WD/ZD!AP5:AQ85").execute()

# создаем новый excel-файл
wb = openpyxl.Workbook()

# добавляем новый лист
wb.create_sheet(title='Инвентаризация', index=0)

# получаем лист, с которым будем работать
sheet = wb['Инвентаризация']

#  Подгоняем по ширине столбцы
sheet.column_dimensions['A'].width = 65
sheet.column_dimensions['B'].width = 11
sheet.column_dimensions['C'].width = 14
sheet.column_dimensions['D'].width = 5

#  Запись хозтоваров
sheet.cell(row=1, column=1).value = 'Список хозтоваров:'
sheet.cell(row=1, column=1).font = Font(size=25)

count = 3

for row in range(count, len(order_house)+count):
    for col in range(1, 5):
        if col == 4:
            sheet.cell(row=row, column=col).value = int(order_house[row - 3][col - 1])
        else:
            sheet.cell(row=row, column=col).value = order_house[row-3][col-1]
        if col > 1:
            sheet.cell(row=row, column=col).alignment = Alignment(horizontal='center')

count += len(order_house) + 2


#  Запись чаёв
sheet.cell(row=count, column=1).value = 'Список чаёв:'
sheet.cell(row=count, column=1).font = Font(size=25)
count += 2

for row in range(count, len(order_tea)+count):
    for col in range(1, 3):
        if col == 2:
            sheet.cell(row=row, column=col).value = int(order_tea[row - count][col - 1])
            sheet.cell(row=row, column=col).alignment = Alignment(horizontal='center')
        else:
            sheet.cell(row=row, column=col).value = order_tea[row-count][col-1]


#  Сохранение в файл
wb.save('Order.xlsx')

#  Открытие файла (файл должен быть закрыт)
os.startfile('Order.xlsx')