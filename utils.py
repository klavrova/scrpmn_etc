# Несколько процедур, которые используются в различных скриптах этого проекта


def find_end(workbook):
    e = 1
    while workbook[f'd{e}'].value is not None:
        e += 1
    return e


def convert_to_number(s):
    s = s.strip()
    try:
        return int(s.replace(',00', ''))
    except ValueError:
        try:
            return float(s.replace(',', '.'))
        except ValueError:
            return s
