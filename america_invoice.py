# Парсит и разбивает по столбцам данные из счетов американских товаров

import re
import os
import argparse
from contextlib import suppress

from openpyxl import load_workbook

ART_ARRAY = [27, 12, 14, 20]


def remove_article_beginning(cell, data):
    for s in data:
        if cell.startswith(str(s)):
            return cell.replace(str(s), '', 1)
    return cell


def get_input_filename():
    parser = argparse.ArgumentParser(description='Программа для конвертирования счетов америки в excel')
    parser.add_argument('path', metavar='P', type=str, help='Путь к таблице')
    return parser.parse_args().path


def parse_america_invoice(input_file):
    sheet = load_workbook(input_file)
    wb = sheet[sheet.get_sheet_names()[0]]
    wb['A1'].value = 'АРТИКУЛ'
    wb['B1'].value = 'НАЗВАНИЕ'
    wb['C1'].value = 'РРЦ'
    wb['D1'].value = 'КОД'
    wb['E1'].value = 'К-ВО'
    wb['F1'].value = 'ЦЕНА'
    wb['G1'].value = 'СУММА'
    cell = 2
    while wb[f'A{cell}'].value is not None:
        new_art_array = [f'{s} ' for s in ART_ARRAY]
        wb[f'A{cell}'].value = remove_article_beginning(wb[f'A{cell}'].value, new_art_array)
        wb[f'A{cell}'].value = re.sub(' c$', '', str(wb[f'A{cell}'].value))
        data_list = wb[f'A{cell}'].value.split(' ')
        wb[f'G{cell}'].value = float(data_list.pop())
        wb[f'F{cell}'].value = float(data_list.pop())
        wb[f'E{cell}'].value = int(data_list.pop().replace('.00', ''))
        wb[f'D{cell}'].value = data_list.pop()
        wb[f'C{cell}'].value = float(data_list.pop())
        with suppress(ValueError):
            data_list[0] = int(data_list[0])
        if type(data_list[0]) is str:
            data_list[0] = remove_article_beginning(data_list[0], ART_ARRAY)
        wb[f'A{cell}'].value = data_list.pop(0)
        wb[f'B{cell}' + str(cell)].value = ' '.join(data_list)
        cell += 1
    sheet.save(f'{os.path.splitext(input_file)[0]}(1).xlsx')


if __name__ == '__main__':
    parse_america_invoice(get_input_filename())
