# Скрипт для парсинга счетов от поставщика 7 Dots

import itertools
import string

import utils

from openpyxl import load_workbook


def parse_7dots_invoice(input_filename, output_filename):
    sheet = load_workbook(input_filename)
    wb = sheet[sheet.get_sheet_names()[0]]
    wb['A1'].value = 'No.'
    wb['B1'].value = 'Product name'
    wb['C1'].value = 'Unit'
    wb['D1'].value = 'Quantity'
    wb['E1'].value = 'Unit Net price'
    wb['F1'].value = 'Discount'
    wb['G1'].value = 'Total Net price'
    wb['H1'].value = 'VAT'
    wb['I1'].value = 'Gross'
    cell = 2
    end = utils.find_end(wb)
    while wb[f'A{cell}'].value is not None:
        material = [utils.convert_to_number(x) for x in wb[f'A{cell}'].value.split('|')][1:]
        for idx, row in enumerate(itertools.chain(string.ascii_uppercase[:end])):
            wb[f'{row}{cell}'].value = material[idx]
        cell += 1
    sheet.save(output_filename)


if __name__ == '__main__':
    # скрипт писался для одного случая, поэтому жестко забито название файла
    parse_7dots_invoice('C:\РАБОЧАЯ\\7dot_invoice.xlsx', 'C:\РАБОЧАЯ\\7dot_invoice Ready.xlsx')
