from openpyxl import load_workbook


def load(filename, sheetname):
    wb = load_workbook(filename=filename)
    sheet = wb[sheetname]

    row_value = []  # 列のデータ
    sheet_value = []  # シートのデータ

    for row in sheet:
        for cell in row:
            row_value.append(cell.value)
        sheet_value.append(row_value)
        row_value = []

    del sheet_value[0]  # 先頭行を削除
    # print(sheet_value)

    wb.close()
    return sheet_value
