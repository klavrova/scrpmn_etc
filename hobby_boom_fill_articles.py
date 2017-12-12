# Скрипт для дополнения информации на выгрузку от поставщика Hobby Boom

import sys
import os

import utils

from openpyxl import load_workbook


def fill_hobby_boom_articles():
    filename = sys.argv[1:][0]
    sheet = load_workbook(filename)
    wb = sheet[sheet.get_sheet_names()[0]]
    for cell in range(2, utils.find_end(wb)):
        article = wb[f'O{cell}'].value
        wb[f'S{cell}'].value = article
        wb[f'S{cell}'].hyperlink = f'http://www.hobby-opt.ru/files/originals/{article}.jpg'
    sheet.save(f'{os.path.splitext(filename)[0]}(1).xlsx')


if __name__ == '__main__':
    fill_hobby_boom_articles()
