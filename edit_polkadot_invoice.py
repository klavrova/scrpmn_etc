# Скрипт для парсинга счетов от поставщика Polkadot

import sys
import os

import utils

from openpyxl import load_workbook


def polkadot(filename):
    sheet = load_workbook(filename)
    wb = sheet[sheet.get_sheet_names()[0]]
    wb['A1'] = '№'
    wb['B1'] = 'Артикул'
    wb['C1'] = 'Товары'
    wb['D1'] = 'Кол-во'
    wb['E1'] = 'Ед'
    wb['F1'] = 'Цена'
    wb['G1'] = 'Сумма'
    cell = 2
    while wb[f'A{cell}'].value is not None:
        parsed = wb[f'A{cell}'].value.split(' ')
        wb[f'A{cell}'].value = int(parsed.pop(0))
        wb[f'B{cell}'].value = parsed.pop(0)
        wb[f'F{cell}'].value = utils.convert_to_number(parsed.pop())
        wb[f'E{cell}'].value = parsed.pop()
        wb[f'D{cell}'].value = float(parsed.pop().replace(',', '.'))
        wb[f'C{cell}'].value = ' '.join(parsed)
        wb[f'G{cell}'].value = wb[f'D{cell}'].value * wb[f'F{cell}'].value
        cell += 1
    sheet.save(f'{os.path.splitext(filename)[0]}(1).xlsx')


if __name__ == '__main__':
    polkadot(sys.argv[1])
