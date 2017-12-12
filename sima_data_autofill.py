# Скрипт для дополнения информации на выгрузку от поставщика SimaLand

import sys
import math
import os

import utils

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook


def fill_data(filename):
    sheet = load_workbook(filename)
    wb = sheet[sheet.get_sheet_names()[0]]
    for cell in range(2, utils.find_end(wb)):
        if wb[f'b{cell}'].value is not None:
            article = wb[f'o{cell}'].value
            link = f'https://www.sima-land.ru/{article}/'
            wb[f'f{cell}'].hyperlink = link
            wb[f's{cell}'].hyperlink = link
            description = BeautifulSoup(requests.get(link).content, 'html.parser').body.select('.b-properties__item')
            for e in description:
                if 'Вес' in e.select('.b-properties__label')[0].get_text():
                    weight = e.select('.b-properties__label')[0].get_text().replace(' г', '')
                    try:
                        weight = int(weight)
                    except ValueError:
                        weight = math.ceil(float(weight))
                    wb[f'k{cell}'].value = weight + 2
            wb[f'b{cell}'].value = (wb[f'b{cell}'].value.replace('*', 'x')
                                    .replace(' x ', 'x').replace('см', ' см').replace('  ', ' '))
            parts = wb[f'b{cell}'].value.split(' ')
            if parts[-1] == str(article):
                parts = parts.pop()
            if parts[-1] == 'см':
                wb[f'h{cell}'].value = ' '.join(parts[-2:])
                wb[f'b{cell}'].value = ' '.join(parts[:-2]).rstrip(',')
    sheet.save(f'{os.path.splitext(filename)[0]}(1).xlsx')


if __name__ == '__main__':
    fill_data(sys.argv[1:][0])
